export function getSquirrelShortcutCommand(event: string, executableName: string, updateExecutable: string) {
  if (event === '--squirrel-install' || event === '--squirrel-updated') {
    return {
      executable: updateExecutable,
      args: [`--createShortcut=${executableName}`, '--shortcut-locations=Desktop,StartMenu'],
    };
  }
  if (event === '--squirrel-uninstall') {
    return { executable: updateExecutable, args: [`--removeShortcut=${executableName}`] };
  }
  return null;
}

export async function cleanupConferenceTaskOnUninstall(
  event: string,
  expectedUpdateExe: string,
  gateway: { removeIfOwned(expectedUpdateExe: string): Promise<'removed' | 'skipped'> },
): Promise<'removed' | 'skipped' | 'untouched'> {
  if (event !== '--squirrel-uninstall') return 'untouched';
  return gateway.removeIfOwned(expectedUpdateExe);
}
