// Subscribe only after recovery and the editor are ready. Retain a file read
// during a modal/busy transition instead of dropping it through runBusy().
export function installRequestedFileOpener({ desktop, canOpen, openFile, onError }) {
  let requested = true, pending, running = false, disposed = false, timer;
  function schedule() {
    if (!disposed && !running && !timer) timer = setTimeout(() => {
      timer = undefined;
      void drain();
    }, 75);
  }
  async function drain() {
    running = true;
    try {
      while (!disposed && (requested || pending)) {
        if (!canOpen()) break;
        if (!pending) {
          requested = false;
          try { pending = await desktop.takeRequestedFile(); }
          catch (error) {
            // The main process notifies us if more paths remain. A persistent
            // IPC failure must wait for a new notification, not retry forever.
            onError(error);
            break;
          }
          if (!pending) continue;
        }
        if (!canOpen()) break;
        const file = pending;
        pending = undefined;
        requested = true;
        try { await openFile(file); }
        catch (error) { onError(error); }
      }
    } finally {
      running = false;
      if (requested || pending) schedule();
    }
  }
  const unsubscribe = desktop.onOpenRequested(() => { requested = true; schedule(); });
  schedule();
  return () => { disposed = true; clearTimeout(timer); unsubscribe(); };
}
