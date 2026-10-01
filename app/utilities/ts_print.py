import builtins
import time

_original_print = builtins.print

def _timestamped_print(*args, **kwargs):
    ts = time.strftime("[%H:%M:%S]")
    # Handle multi-line strings so every line gets its own timestamp,
    # not just the first line of a multi-line print() call.
    sep = kwargs.get("sep", " ")
    text = sep.join(str(a) for a in args)
    lines = text.split("\n")
    for line in lines:
        _original_print(ts, line, **{k: v for k, v in kwargs.items() if k != "sep"})

builtins.print = _timestamped_print