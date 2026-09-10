"""Background work with main-thread callbacks and stale-result rejection."""
import queue
import threading


class TaskRunner:
    def __init__(self, root):
        self.root = root
        self.events = queue.Queue()
        self.tokens = {}
        self.busy = set()
        self.closed = False
        self.timer = root.after(50, self._drain)

    def invalidate(self, key=None):
        for item in list(self.tokens) if key is None else [key]:
            self.tokens[item] = object()
            self.busy.discard(item)

    def call_soon(self, callback):
        if not self.closed:
            self.events.put((None, None, lambda _: callback(), None))

    def submit(self, key, work, done, failed, replace=False):
        if self.closed or (key in self.busy and not replace):
            return False
        token = object()
        self.tokens[key] = token
        self.busy.add(key)

        def worker():
            try:
                self.events.put((key, token, done, work()))
            except Exception as exc:
                self.events.put((key, token, failed, str(exc)))

        threading.Thread(target=worker, daemon=True).start()
        return True

    def _drain(self):
        if self.closed:
            return
        try:
            while True:
                key, token, callback, result = self.events.get_nowait()
                if key is None or self.tokens.get(key) is token:
                    self.busy.discard(key)
                    try:
                        callback(result)
                    except Exception:
                        import sys
                        self.root.report_callback_exception(*sys.exc_info())
        except queue.Empty:
            pass
        self.timer = self.root.after(50, self._drain)

    def close(self):
        self.closed = True
        self.invalidate()
        self.root.after_cancel(self.timer)
