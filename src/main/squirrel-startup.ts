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
