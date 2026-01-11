// Utility functions for the Library System
export function generateId() {
  return '_' + Math.random().toString(36).substr(2, 9);
}

export function formatDate(date) {
  return date.toISOString().split('T')[0];
}
