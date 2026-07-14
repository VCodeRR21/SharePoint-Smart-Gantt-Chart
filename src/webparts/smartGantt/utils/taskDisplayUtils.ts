import { ITask } from '../models';
import { parseDateOnly, todayLocalMidnight } from './dateUtils';

/** True when a task's due date has passed and it isn't Completed/Cancelled. */
export function isOverdue(task: ITask): boolean {
  if (!task.dueDate || task.status === 'Completed' || task.status === 'Cancelled') return false;
  const due = parseDateOnly(task.dueDate);
  return !!due && due < todayLocalMidnight();
}

/** Up to two initials from a display name, for avatar circles. */
export function initials(name: string): string {
  if (!name) return '?';
  return name.split(' ').map(p => p[0]).join('').substring(0, 2).toUpperCase();
}

/** Deterministic HSL color from a name, so the same person's avatar always matches. */
export function stringToColor(s: string): string {
  let hash = 0;
  for (let i = 0; i < s.length; i++) hash = s.charCodeAt(i) + ((hash << 5) - hash);
  return `hsl(${Math.abs(hash) % 360}, 55%, 45%)`;
}
