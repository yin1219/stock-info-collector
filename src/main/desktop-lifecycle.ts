interface WindowPort {
  on(event: 'close', listener: (event: { preventDefault(): void }) => void): void;
  show(): void;
  hide(): void;
  focus(): void;
  isDestroyed(): boolean;
}

interface TrayPort {
  on(event: 'click', listener: () => void): void;
}

interface MenuActions {
  open(): void;
  quit(): void;
}

export function installTrayLifecycle(dependencies: {
  window: WindowPort;
  tray: TrayPort;
  menu: { setActions(actions: MenuActions): void };
  stopMonitoring(): void;
  quitApp(): void;
}): void {
  let quitting = false;
  dependencies.window.on('close', (event) => {
    if (quitting) return;
    event.preventDefault();
    dependencies.window.hide();
  });
  const open = () => {
    if (dependencies.window.isDestroyed()) return;
    dependencies.window.show();
    dependencies.window.focus();
  };
  dependencies.tray.on('click', open);
  dependencies.menu.setActions({
    open,
    quit() {
      if (quitting) return;
      quitting = true;
      dependencies.stopMonitoring();
      dependencies.quitApp();
    },
  });
}
