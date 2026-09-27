export function openStorageBeforeServices<T extends { close?: () => void }>(
  openAndMigrate: () => T,
  startServices: (database: T) => void,
): T {
  const database = openAndMigrate();
  try {
    startServices(database);
    return database;
  } catch (error) {
    database.close?.();
    throw error;
  }
}
