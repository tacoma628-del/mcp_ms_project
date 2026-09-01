"""
MCP Server for Microsoft Project (.mpp) files using mpxj (org.mpxj).
Supports reading AND writing: tasks, resources, assignments, project properties.

Write operations save to MSPDI XML format (.xml) which MS Project opens natively.
Native .mpp write is not supported by mpxj — open the saved .xml in MS Project
and use File > Save As to convert back to .mpp if needed.
"""

import json
import os
import jpype
import mpxj  # adds JARs to classpath
from pathlib import Path
from datetime import datetime
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("msproject")

# ---------------------------------------------------------------------------
# JVM / class helpers
# ---------------------------------------------------------------------------

def _ensure_jvm():
    if not jpype.isJVMStarted():
        jpype.startJVM()


def _cls(name: str):
    _ensure_jvm()
    return jpype.JClass(name)


def _reader():
    return _cls("org.mpxj.reader.UniversalProjectReader")


def _mspdi_writer():
    return _cls("org.mpxj.mspdi.MSPDIWriter")


def _Duration():
    return _cls("org.mpxj.Duration")


def _TimeUnit():
    return _cls("org.mpxj.TimeUnit")


def _ResourceType():
    return _cls("org.mpxj.ResourceType")


def _LocalDateTime():
    return _cls("java.time.LocalDateTime")


def _BigDecimal(value: float):
    return _cls("java.math.BigDecimal")(str(value))


# ---------------------------------------------------------------------------
# Internal utilities
# ---------------------------------------------------------------------------

def _format_date(d):
    if d is None:
        return None
    try:
        return f"{d.getYear():04d}-{d.getMonthValue():02d}-{d.getDayOfMonth():02d}"
    except Exception:
        return str(d)


def _format_duration(dur):
    if dur is None:
        return None
    try:
        return f"{dur.getDuration()} {dur.getUnits()}"
    except Exception:
        return str(dur)


def _parse_date(date_str: str):
    """Parse ISO date string (YYYY-MM-DD) to Java LocalDateTime at 08:00."""
    dt = datetime.fromisoformat(date_str)
    return _LocalDateTime().of(dt.year, dt.month, dt.day, 8, 0)


def _load_project(file_path: str):
    """
    Load a project file and re-enable auto-numbering.
    UniversalProjectReader turns auto-ID assignment OFF on load (to preserve
    the file's existing numbering), which means any addTask()/addResource()
    afterwards gets a null ID and blows up downstream. Turn it back on so
    write operations against an existing file behave the same as against a
    freshly-created one.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    project = _reader()().read(str(path))
    if project is None:
        raise ValueError(f"Could not parse project file (unsupported or corrupt): {file_path}")
    cfg = project.getProjectConfig()
    cfg.setAutoTaskID(True)
    cfg.setAutoTaskUniqueID(True)
    cfg.setAutoResourceID(True)
    cfg.setAutoResourceUniqueID(True)
    cfg.setAutoAssignmentUniqueID(True)
    cfg.setAutoRelationUniqueID(True)
    cfg.setAutoCalendarUniqueID(True)
    cfg.setAutoOutlineLevel(True)
    cfg.setAutoOutlineNumber(True)
    cfg.setAutoWBS(True)
    return project


def _default_output(input_path: str, output_path: str | None) -> str:
    """
    Determine output path. Write always produces MSPDI XML.
    If input is .mpp and no output_path given, save next to original as .xml.
    """
    if output_path:
        p = Path(output_path)
        if p.suffix.lower() == ".mpp":
            p = p.with_suffix(".xml")
        return str(p)
    p = Path(input_path)
    if p.suffix.lower() == ".mpp":
        return str(p.with_suffix(".xml"))
    return str(p)


def _save(project, output_path: str) -> str:
    """
    Save a project, fixing up hierarchy/order first.
    MSPDI encodes parent/child structure by task ORDER in the file plus outline
    level, not an explicit parent reference. addTask(parent) on a project that
    was loaded from disk appends the new child at the end of the task list
    instead of positioning it after the parent's other descendants, which
    silently corrupts the hierarchy on the next read (a later sibling task can
    end up misread as the summary task instead of the real parent). Calling
    synchronizeTaskIDToHierarchy() regroups tasks into the correct order before
    every write, so incremental edits across separate load/save calls (which is
    how every MCP tool call here works) don't accumulate this corruption.
    """
    project.getTasks().synchronizeTaskIDToHierarchy()
    project.updateStructure()
    _mspdi_writer()().write(project, output_path)
    return output_path


# ---------------------------------------------------------------------------
# READ tools
# ---------------------------------------------------------------------------

@mcp.tool()
def get_project_summary(file_path: str) -> str:
    """
    Get a summary of an MS Project file: title, author, dates, task/resource counts.

    Args:
        file_path: Full path to the .mpp (or .xml, .mpx, .xer) file.
    """
    try:
        project = _load_project(file_path)
        props = project.getProjectProperties()
        tasks = [t for t in project.getTasks() if t.getName() is not None]
        resources = [r for r in project.getResources() if r.getName() is not None]
        return json.dumps({
            "title": str(props.getProjectTitle()) if props.getProjectTitle() else Path(file_path).stem,
            "author": str(props.getAuthor()) if props.getAuthor() else None,
            "company": str(props.getCompany()) if props.getCompany() else None,
            "start_date": _format_date(props.getStartDate()),
            "finish_date": _format_date(props.getFinishDate()),
            "status_date": _format_date(props.getStatusDate()),
            "task_count": len(tasks),
            "resource_count": len(resources),
            "currency_symbol": str(props.getCurrencySymbol()) if props.getCurrencySymbol() else None,
        }, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_tasks(file_path: str, include_summary_tasks: bool = True, max_tasks: int = 200) -> str:
    """
    Get all tasks with key fields: dates, % complete, WBS, predecessors, cost.

    Args:
        file_path: Full path to the project file.
        include_summary_tasks: Include parent/summary tasks (default True).
        max_tasks: Max tasks to return (default 200).
    """
    try:
        project = _load_project(file_path)
        result = []
        for task in project.getTasks():
            if task.getName() is None:
                continue
            if not include_summary_tasks and task.getSummary():
                continue
            result.append({
                "id": task.getID(),
                "unique_id": task.getUniqueID(),
                "name": str(task.getName()),
                "outline_level": task.getOutlineLevel(),
                "outline_number": str(task.getOutlineNumber()) if task.getOutlineNumber() else None,
                "wbs": str(task.getWBS()) if task.getWBS() else None,
                "is_summary": bool(task.getSummary()),
                "is_milestone": bool(task.getMilestone()),
                "percent_complete": float(task.getPercentageComplete()) if task.getPercentageComplete() is not None else None,
                "start": _format_date(task.getStart()),
                "finish": _format_date(task.getFinish()),
                "baseline_start": _format_date(task.getBaselineStart()),
                "baseline_finish": _format_date(task.getBaselineFinish()),
                "duration": _format_duration(task.getDuration()),
                "work": _format_duration(task.getWork()),
                "cost": float(task.getCost()) if task.getCost() is not None else None,
                "priority": str(task.getPriority()) if task.getPriority() else None,
                "notes": str(task.getNotes()) if task.getNotes() else None,
                "predecessors": str(task.getPredecessors()) if task.getPredecessors() else None,
            })
            if len(result) >= max_tasks:
                break
        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_resources(file_path: str) -> str:
    """
    Get all resources (people, equipment, materials) with rates and units.

    Args:
        file_path: Full path to the project file.
    """
    try:
        project = _load_project(file_path)
        result = []
        for r in project.getResources():
            if r.getName() is None:
                continue
            result.append({
                "id": r.getID(),
                "unique_id": r.getUniqueID(),
                "name": str(r.getName()),
                "type": str(r.getType()) if r.getType() else None,
                "email": str(r.getEmailAddress()) if r.getEmailAddress() else None,
                "max_units": float(r.getMaxUnits()) if r.getMaxUnits() is not None else None,
                "standard_rate": str(r.getStandardRate()) if r.getStandardRate() else None,
                "overtime_rate": str(r.getOvertimeRate()) if r.getOvertimeRate() else None,
                "group": str(r.getGroup()) if r.getGroup() else None,
                "notes": str(r.getNotes()) if r.getNotes() else None,
            })
        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_assignments(file_path: str) -> str:
    """
    Get all resource-to-task assignments with work and cost.

    Args:
        file_path: Full path to the project file.
    """
    try:
        project = _load_project(file_path)
        result = []
        for a in project.getResourceAssignments():
            task = a.getTask()
            resource = a.getResource()
            result.append({
                "task_id": task.getID() if task else None,
                "task_name": str(task.getName()) if task else None,
                "resource_id": resource.getID() if resource else None,
                "resource_name": str(resource.getName()) if resource else None,
                "units": float(a.getUnits()) if a.getUnits() is not None else None,
                "work": _format_duration(a.getWork()),
                "actual_work": _format_duration(a.getActualWork()),
                "start": _format_date(a.getStart()),
                "finish": _format_date(a.getFinish()),
                "cost": float(a.getCost()) if a.getCost() is not None else None,
            })
        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_critical_path(file_path: str) -> str:
    """
    Get all tasks on the critical path.

    Args:
        file_path: Full path to the project file.
    """
    try:
        project = _load_project(file_path)
        result = [
            {
                "id": t.getID(),
                "name": str(t.getName()),
                "start": _format_date(t.getStart()),
                "finish": _format_date(t.getFinish()),
                "duration": _format_duration(t.getDuration()),
                "percent_complete": float(t.getPercentageComplete()) if t.getPercentageComplete() is not None else None,
                "is_milestone": bool(t.getMilestone()),
            }
            for t in project.getTasks()
            if t.getName() is not None and t.getCritical()
        ]
        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_overdue_tasks(file_path: str, status_date: str = None) -> str:
    """
    Get tasks past their finish date that are not 100% complete.

    Args:
        file_path: Full path to the project file.
        status_date: Reference date YYYY-MM-DD (default: today).
    """
    try:
        project = _load_project(file_path)
        ref = datetime.fromisoformat(status_date) if status_date else datetime.now()
        result = []
        for task in project.getTasks():
            if task.getName() is None or task.getSummary():
                continue
            finish = task.getFinish()
            if finish is None:
                continue
            try:
                finish_py = datetime(finish.getYear(), finish.getMonthValue(), finish.getDayOfMonth())
            except Exception:
                continue
            pct = float(task.getPercentageComplete()) if task.getPercentageComplete() is not None else 0.0
            if finish_py < ref and pct < 100.0:
                result.append({
                    "id": task.getID(),
                    "name": str(task.getName()),
                    "finish": _format_date(task.getFinish()),
                    "percent_complete": pct,
                    "is_milestone": bool(task.getMilestone()),
                })
        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# WRITE tools
# ---------------------------------------------------------------------------

@mcp.tool()
def add_task(
    file_path: str,
    name: str,
    start: str = None,
    finish: str = None,
    duration_days: float = None,
    percent_complete: float = None,
    milestone: bool = False,
    notes: str = None,
    parent_task_id: int = None,
    output_path: str = None,
) -> str:
    """
    Add a new task to the project and save.
    Saves as MSPDI XML (.xml) — open in MS Project to save back as .mpp.

    Args:
        file_path: Source project file.
        name: Task name.
        start: Start date YYYY-MM-DD.
        finish: Finish date YYYY-MM-DD.
        duration_days: Duration in working days (ignored if both start+finish given).
        percent_complete: 0-100.
        milestone: True to mark as milestone.
        notes: Task notes/description.
        parent_task_id: ID of the parent task (creates a subtask).
        output_path: Where to save the result (default: same dir, .xml extension).
    """
    try:
        project = _load_project(file_path)

        if parent_task_id is not None:
            parent = next((t for t in project.getTasks() if t.getUniqueID() == parent_task_id), None)
            if parent is None:
                return json.dumps({"error": f"Parent task ID {parent_task_id} not found"})
            task = parent.addTask()
        else:
            task = project.addTask()

        task.setName(name)

        if start:
            task.setStart(_parse_date(start))
        if finish:
            task.setFinish(_parse_date(finish))
        if duration_days is not None:
            task.setDuration(_Duration().getInstance(duration_days, _TimeUnit().DAYS))
        if percent_complete is not None:
            task.setPercentageComplete(_BigDecimal(percent_complete))
        if milestone:
            task.setMilestone(True)
        if notes:
            task.setNotes(notes)

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({
            "success": True,
            "task_id": task.getUniqueID(),
            "task_name": str(task.getName()),
            "saved_to": out,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def update_task(
    file_path: str,
    task_id: int,
    name: str = None,
    start: str = None,
    finish: str = None,
    duration_days: float = None,
    percent_complete: float = None,
    milestone: bool = None,
    notes: str = None,
    output_path: str = None,
) -> str:
    """
    Update fields of an existing task and save.

    Args:
        file_path: Source project file.
        task_id: ID of the task to update (use get_tasks to find IDs).
        name: New task name (optional).
        start: New start date YYYY-MM-DD (optional).
        finish: New finish date YYYY-MM-DD (optional).
        duration_days: New duration in working days (optional).
        percent_complete: New % complete 0-100 (optional).
        milestone: Set/unset milestone flag (optional).
        notes: New notes (optional).
        output_path: Where to save the result.
    """
    try:
        project = _load_project(file_path)
        task = next((t for t in project.getTasks() if t.getUniqueID() == task_id), None)
        if task is None:
            return json.dumps({"error": f"Task ID {task_id} not found"})

        if name is not None:
            task.setName(name)
        if start is not None:
            task.setStart(_parse_date(start))
        if finish is not None:
            task.setFinish(_parse_date(finish))
        if duration_days is not None:
            task.setDuration(_Duration().getInstance(duration_days, _TimeUnit().DAYS))
        if percent_complete is not None:
            task.setPercentageComplete(_BigDecimal(percent_complete))
        if milestone is not None:
            task.setMilestone(milestone)
        if notes is not None:
            task.setNotes(notes)

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({
            "success": True,
            "task_id": task_id,
            "task_name": str(task.getName()),
            "saved_to": out,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def delete_task(file_path: str, task_id: int, output_path: str = None) -> str:
    """
    Delete a task from the project and save.

    Args:
        file_path: Source project file.
        task_id: ID of the task to delete.
        output_path: Where to save the result.
    """
    try:
        project = _load_project(file_path)
        task = next((t for t in project.getTasks() if t.getUniqueID() == task_id), None)
        if task is None:
            return json.dumps({"error": f"Task ID {task_id} not found"})

        task_name = str(task.getName())
        task.remove()

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({"success": True, "deleted_task": task_name, "saved_to": out})
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def add_dependency(
    file_path: str,
    predecessor_task_id: int,
    successor_task_id: int,
    dependency_type: str = "FINISH_START",
    lag_days: float = 0,
    output_path: str = None,
) -> str:
    """
    Link two tasks with a predecessor/successor dependency and save.
    Note: this stores the relationship but does not itself recompute dates or
    the critical path -- open the saved .xml in MS Project / Project Plan 365
    to let it schedule the tasks, then re-read the file to see the result.

    Args:
        file_path: Source project file.
        predecessor_task_id: ID of the task that must happen first.
        successor_task_id: ID of the task that depends on it.
        dependency_type: One of FINISH_START, START_START, FINISH_FINISH, START_FINISH.
        lag_days: Lag (or negative lead) in working days (default 0).
        output_path: Where to save the result.
    """
    try:
        project = _load_project(file_path)

        predecessor = next((t for t in project.getTasks() if t.getUniqueID() == predecessor_task_id), None)
        if predecessor is None:
            return json.dumps({"error": f"Predecessor task ID {predecessor_task_id} not found"})

        successor = next((t for t in project.getTasks() if t.getUniqueID() == successor_task_id), None)
        if successor is None:
            return json.dumps({"error": f"Successor task ID {successor_task_id} not found"})

        RelationType = _cls("org.mpxj.RelationType")
        rt = getattr(RelationType, dependency_type.upper(), RelationType.FINISH_START)

        Relation = _cls("org.mpxj.Relation")
        builder = Relation.Builder().predecessorTask(predecessor).successorTask(successor).type(rt)
        if lag_days:
            builder = builder.lag(_Duration().getInstance(lag_days, _TimeUnit().DAYS))
        successor.addPredecessor(builder)

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({
            "success": True,
            "predecessor": str(predecessor.getName()),
            "successor": str(successor.getName()),
            "type": dependency_type.upper(),
            "saved_to": out,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def remove_dependency(
    file_path: str,
    predecessor_task_id: int,
    successor_task_id: int,
    dependency_type: str = "FINISH_START",
    output_path: str = None,
) -> str:
    """
    Remove a predecessor/successor link between two tasks and save.
    Useful when restructuring a schedule -- e.g. a task became a summary task
    with subtasks, and the old direct link needs to move to a child task instead.

    Args:
        file_path: Source project file.
        predecessor_task_id: ID of the predecessor task in the existing link.
        successor_task_id: ID of the successor task in the existing link.
        dependency_type: The relation type to remove (must match what was added).
        output_path: Where to save the result.
    """
    try:
        project = _load_project(file_path)

        predecessor = next((t for t in project.getTasks() if t.getUniqueID() == predecessor_task_id), None)
        if predecessor is None:
            return json.dumps({"error": f"Predecessor task ID {predecessor_task_id} not found"})

        successor = next((t for t in project.getTasks() if t.getUniqueID() == successor_task_id), None)
        if successor is None:
            return json.dumps({"error": f"Successor task ID {successor_task_id} not found"})

        RelationType = _cls("org.mpxj.RelationType")
        rt = getattr(RelationType, dependency_type.upper(), RelationType.FINISH_START)

        removed = successor.removePredecessor(predecessor, rt, None)
        if not removed:
            return json.dumps({"error": f"No {dependency_type.upper()} link found from {predecessor_task_id} to {successor_task_id}"})

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({
            "success": True,
            "predecessor": str(predecessor.getName()),
            "successor": str(successor.getName()),
            "saved_to": out,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def add_resource(
    file_path: str,
    name: str,
    resource_type: str = "WORK",
    email: str = None,
    max_units: float = 1.0,
    standard_rate: float = None,
    group: str = None,
    notes: str = None,
    output_path: str = None,
) -> str:
    """
    Add a new resource (person, equipment, or material) and save.

    Args:
        file_path: Source project file.
        name: Resource name.
        resource_type: WORK, MATERIAL, or COST (default: WORK).
        email: Email address.
        max_units: Max allocation units (1.0 = 100%, default 1.0).
        standard_rate: Hourly/daily rate (numeric).
        group: Group/department name.
        notes: Resource notes.
        output_path: Where to save the result.
    """
    try:
        project = _load_project(file_path)
        resource = project.addResource()
        resource.setName(name)

        rt_map = {"WORK": "WORK", "MATERIAL": "MATERIAL", "COST": "COST"}
        rt_str = rt_map.get(resource_type.upper(), "WORK")
        ResourceType = _ResourceType()
        resource.setType(getattr(ResourceType, rt_str))

        if email:
            resource.setEmailAddress(email)
        if max_units is not None:
            resource.setMaxUnits(max_units)
        if group:
            resource.setGroup(group)
        if notes:
            resource.setNotes(notes)

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({
            "success": True,
            "resource_id": resource.getUniqueID(),
            "resource_name": str(resource.getName()),
            "saved_to": out,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def update_resource(
    file_path: str,
    resource_id: int,
    name: str = None,
    email: str = None,
    max_units: float = None,
    group: str = None,
    notes: str = None,
    output_path: str = None,
) -> str:
    """
    Update fields of an existing resource and save.

    Args:
        file_path: Source project file.
        resource_id: ID of the resource to update.
        name: New name (optional).
        email: New email (optional).
        max_units: New max units 0-1 (optional).
        group: New group (optional).
        notes: New notes (optional).
        output_path: Where to save the result.
    """
    try:
        project = _load_project(file_path)
        resource = next((r for r in project.getResources() if r.getUniqueID() == resource_id), None)
        if resource is None:
            return json.dumps({"error": f"Resource ID {resource_id} not found"})

        if name is not None:
            resource.setName(name)
        if email is not None:
            resource.setEmailAddress(email)
        if max_units is not None:
            resource.setMaxUnits(max_units)
        if group is not None:
            resource.setGroup(group)
        if notes is not None:
            resource.setNotes(notes)

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({
            "success": True,
            "resource_id": resource_id,
            "resource_name": str(resource.getName()),
            "saved_to": out,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def delete_resource(file_path: str, resource_id: int, output_path: str = None) -> str:
    """
    Delete a resource from the project and save.
    Also removes all assignments for this resource.

    Args:
        file_path: Source project file.
        resource_id: ID of the resource to delete.
        output_path: Where to save the result.
    """
    try:
        project = _load_project(file_path)
        resource = next((r for r in project.getResources() if r.getUniqueID() == resource_id), None)
        if resource is None:
            return json.dumps({"error": f"Resource ID {resource_id} not found"})

        resource_name = str(resource.getName())
        resource.remove()

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({"success": True, "deleted_resource": resource_name, "saved_to": out})
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def assign_resource(
    file_path: str,
    task_id: int,
    resource_id: int,
    units: float = 1.0,
    output_path: str = None,
) -> str:
    """
    Assign a resource to a task and save.

    Args:
        file_path: Source project file.
        task_id: ID of the task.
        resource_id: ID of the resource.
        units: Allocation units (1.0 = 100%, default 1.0).
        output_path: Where to save the result.
    """
    try:
        project = _load_project(file_path)

        task = next((t for t in project.getTasks() if t.getUniqueID() == task_id), None)
        if task is None:
            return json.dumps({"error": f"Task ID {task_id} not found"})

        resource = next((r for r in project.getResources() if r.getUniqueID() == resource_id), None)
        if resource is None:
            return json.dumps({"error": f"Resource ID {resource_id} not found"})

        assignment = task.addResourceAssignment(resource)
        assignment.setUnits(units)

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({
            "success": True,
            "task_name": str(task.getName()),
            "resource_name": str(resource.getName()),
            "units": units,
            "saved_to": out,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def remove_assignment(
    file_path: str,
    task_id: int,
    resource_id: int,
    output_path: str = None,
) -> str:
    """
    Remove a resource assignment from a task and save.

    Args:
        file_path: Source project file.
        task_id: ID of the task.
        resource_id: ID of the resource to unassign.
        output_path: Where to save the result.
    """
    try:
        project = _load_project(file_path)

        assignment = next(
            (a for a in project.getResourceAssignments()
             if a.getTask() is not None and a.getTask().getUniqueID() == task_id
             and a.getResource() is not None and a.getResource().getUniqueID() == resource_id),
            None
        )
        if assignment is None:
            return json.dumps({"error": f"No assignment found for task {task_id} / resource {resource_id}"})

        task_name = str(assignment.getTask().getName())
        resource_name = str(assignment.getResource().getName())
        assignment.remove()

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({
            "success": True,
            "removed_from_task": task_name,
            "removed_resource": resource_name,
            "saved_to": out,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def update_project_properties(
    file_path: str,
    title: str = None,
    author: str = None,
    company: str = None,
    start_date: str = None,
    status_date: str = None,
    output_path: str = None,
) -> str:
    """
    Update project-level properties (title, author, dates) and save.

    Args:
        file_path: Source project file.
        title: Project title.
        author: Project author/manager name.
        company: Company name.
        start_date: Project start date YYYY-MM-DD.
        status_date: Status/data date YYYY-MM-DD.
        output_path: Where to save the result.
    """
    try:
        project = _load_project(file_path)
        props = project.getProjectProperties()

        if title is not None:
            props.setProjectTitle(title)
        if author is not None:
            props.setAuthor(author)
        if company is not None:
            props.setCompany(company)
        if start_date is not None:
            props.setStartDate(_parse_date(start_date))
        if status_date is not None:
            props.setStatusDate(_parse_date(status_date))

        out = _default_output(file_path, output_path)
        _save(project, out)

        return json.dumps({"success": True, "saved_to": out})
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def convert_project(file_path: str, output_path: str) -> str:
    """
    Convert a project file to a different format.
    Supported output formats: .xml (MSPDI), .mpx, .json, .xer

    Args:
        file_path: Source project file path.
        output_path: Destination path with desired extension.
    """
    try:
        project = _load_project(file_path)
        ext = Path(output_path).suffix.lower()

        if ext == ".xml":
            _cls("org.mpxj.mspdi.MSPDIWriter")().write(project, output_path)
        elif ext == ".mpx":
            _cls("org.mpxj.mpx.MPXWriter")().write(project, output_path)
        elif ext == ".json":
            _cls("org.mpxj.json.JsonWriter")().write(project, output_path)
        elif ext == ".xer":
            _cls("org.mpxj.primavera.PrimaveraXERFileWriter")().write(project, output_path)
        else:
            return json.dumps({"error": f"Unsupported output format: {ext}. Use .xml, .mpx, .json, or .xer"})

        return json.dumps({"success": True, "saved_to": output_path, "size_bytes": os.path.getsize(output_path)})
    except Exception as e:
        return json.dumps({"error": str(e)})


if __name__ == "__main__":
    mcp.run()
