import { WorkdayCalendar } from './workday-calendar';

/**
 * Identifier for a task.
 */
type TaskId = number;

/**
 * Represents a task with scheduling-related fields.
 * This is the core input/output unit in the scheduling system.
 */
export type TaskDefinition = {
  /**
   * Unique identifier of the task.
   */
  id: TaskId;

  /**
   * Start date of the task in "yyyy/MM/dd" format.
   * May be null if not yet determined.
   */
  startDate: string | null;

  /**
   * End date of the task in "yyyy/MM/dd" format.
   * May be null if not yet determined.
   */
  endDate: string | null;

  /**
   * Duration of the task in days.
   * May be null if inferred from startDate and endDate or other constraints.
   */
  duration: number | null;

  /**
   * Set of task IDs that must end before this task can start (startAfterEnd constraint).
   */
  startsAfter: Set<TaskId>;

  /**
   * Set of task IDs that must start after this task ends (endBeforeStart constraint).
   */
  endsBefore: Set<TaskId>;
};

/**
 * Type of dependency between tasks.
 * 'startAfterEnd': this task starts after another ends.
 * 'endBeforeStart': this task ends before another starts.
 */
type DependencyType = 'startAfterEnd' | 'endBeforeStart';

/**
 * Represents a directed dependency between two tasks.
 */
type DependencyEdge = {
  /**
   * The task being depended on.
   */
  from: TaskId;

  /**
   * The task that depends on the "from" task.
   */
  to: TaskId;

  /**
   * The type of dependency between the tasks.
   */
  type: DependencyType;
};

/**
 * A graph structure that contains tasks and their dependency edges.
 */
type TaskGraph = {
  /**
   * A map of task ID to its corresponding TaskDefinition.
   */
  tasks: Map<TaskId, TaskDefinition>;

  /**
   * List of dependency edges between tasks.
   */
  edges: DependencyEdge[];
};

/**
 * Creates a dependency graph from a list of tasks.
 *
 * @param tasks - Array of TaskDefinition objects to include in the graph.
 * @returns A TaskGraph structure representing task dependencies.
 */
function createDependencyGraph(tasks: TaskDefinition[]) {
  const taskGraph: TaskGraph = { tasks: new Map(), edges: [] };
  tasks.forEach(task => {
    taskGraph.tasks.set(task.id, task);
    taskGraph.edges.push(
      ...[
        ...[...task.startsAfter].map(id => ({
          from: id,
          to: task.id,
          type: 'startAfterEnd' as const,
        })),
        ...[...task.endsBefore].map(id => ({
          from: id,
          to: task.id,
          type: 'endBeforeStart' as const,
        })),
      ]
    );
  });
  return taskGraph;
}

/**
 * Performs a topological sort on the given task dependency graph.
 * Throws an error if the graph contains a cycle.
 *
 * @param graph - TaskGraph to be sorted.
 * @returns An array of TaskId sorted in dependency-respecting order.
 * @throws Error if a cycle is found in the task graph.
 */
function topologicalSort(graph: TaskGraph): TaskId[] {
  const inDegree = new Map<TaskId, number>();
  const adjacencyList = new Map<TaskId, TaskId[]>();

  graph.tasks.forEach((_task, id) => {
    inDegree.set(id, 0);
    adjacencyList.set(id, []);
  });

  for (const edge of graph.edges) {
    adjacencyList.get(edge.from)?.push(edge.to);
    inDegree.set(edge.to, (inDegree.get(edge.to) || 0) + 1);
  }

  const queue: TaskId[] = [];
  for (const [id, degree] of inDegree.entries()) {
    if (degree === 0) {
      queue.push(id);
    }
  }

  const sorted: TaskId[] = [];
  while (queue.length > 0) {
    const current = queue.shift()!;
    sorted.push(current);

    for (const neighbor of adjacencyList.get(current)!) {
      inDegree.set(neighbor, inDegree.get(neighbor)! - 1);
      if (inDegree.get(neighbor) === 0) {
        queue.push(neighbor);
      }
    }
  }

  if (sorted.length !== graph.tasks.size) {
    throw new Error('[ScheduleError] Cycle detected in task graph');
  }

  return sorted;
}

/**
 * Calculates the full schedule for the given tasks.
 * This includes computing missing startDate or endDate fields.
 * Does not modify the input array.
 *
 * @param tasks - The list of input tasks with partial information.
 * @param calendar A WorkdayCalendar instance used to compute workday-based date shifts and lookups.
 * @returns A tuple containing a new array of TaskDefinition with startDate/endDate filled in,
 *          and a map of task ID to its corresponding TaskDefinition.
 * @throws Error if scheduling is not possible due to insufficient data or circular dependencies.
 */
export function resolveSchedule(
  tasks: TaskDefinition[],
  calendar: WorkdayCalendar
) {
  const scheduledTasks = tasks.map(task => ({ ...task }));
  const graph = createDependencyGraph(scheduledTasks);
  const sortedIds = topologicalSort(graph);

  sortedIds.forEach(id => {
    const task = graph.tasks.get(id)!;

    if (task.startDate !== null && task.endsBefore.size) {
      throw new Error(
        `[ScheduleError] Invalid schedule for task ${task.id}: both startDate and endsBefore are specified`
      );
    }
    if (task.endDate !== null && task.startsAfter.size) {
      throw new Error(
        `[ScheduleError] Invalid schedule for task ${task.id}: both endDate and startsAfter are specified`
      );
    }
    if (task.startsAfter.size && task.endsBefore.size) {
      throw new Error(
        `[ScheduleError] Invalid schedule for task ${task.id}: both startsAfter and endsBefore are specified`
      );
    }

    if (task.startsAfter.size) {
      const latestEnd = graph.edges
        .filter(e => e.to === id && e.type === 'startAfterEnd')
        .map(e => graph.tasks.get(e.from)!.endDate!)
        .reduce((max, c) => (c > max ? c : max));
      task.startDate = calendar.getNextWorkday(latestEnd);
      if (task.startDate === null) {
        throw new Error(
          `[ScheduleError] Failed to shift start date for task ${task.id}`
        );
      }
    }
    if (task.endsBefore.size) {
      const earliestStart = graph.edges
        .filter(e => e.to === id && e.type === 'endBeforeStart')
        .map(e => graph.tasks.get(e.from)!.startDate!)
        .reduce((min, c) => (c < min ? c : min));
      task.endDate = calendar.getPreviousWorkday(earliestStart);
      if (task.endDate === null) {
        throw new Error(
          `[ScheduleError] Failed to shift end date for task ${task.id}`
        );
      }
    }

    if (task.startDate !== null && task.duration !== null) {
      task.endDate = calendar.getWorkdayAtOffset(
        task.startDate,
        task.duration - 1
      );
      if (task.endDate === null) {
        throw new Error(
          `[ScheduleError] Failed to compute end date for task ${task.id}`
        );
      }
    }
    if (task.endDate !== null && task.duration !== null) {
      task.startDate = calendar.getWorkdayAtOffset(
        task.endDate,
        -(task.duration - 1)
      );
      if (task.startDate === null) {
        throw new Error(
          `[ScheduleError] Failed to compute start date for task ${task.id}`
        );
      }
    }

    if (task.startDate === null)
      throw new Error(
        `[ScheduleError] Failed to compute start date for task ${task.id}: insufficient information`
      );
    if (task.endDate === null)
      throw new Error(
        `[ScheduleError] Failed to compute end date for task ${task.id}: insufficient information`
      );
    if (task.endDate < task.startDate)
      throw new Error(
        `[ScheduleError] Invalid schedule for task ${task.id}: end date (${task.endDate}) is earlier than start date (${task.startDate})`
      );
  });

  return [scheduledTasks, graph.tasks] as const;
}
