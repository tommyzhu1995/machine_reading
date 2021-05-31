import platform
import signal
from functools import wraps


## For linux
def timeout_wrap(interval):
    def decorator(func):
        def handler(signum, frame):
            raise TimeoutError("run {} timeout".format(func.__name__))

        @wraps(func)
        def wrapper(*args, **kwargs):
            signal.signal(signal.SIGALRM, handler)
            signal.alarm(interval)       # interval秒后向进程发送SIGALRM信号
            result = func(*args, **kwargs)
            signal.alarm(0)              # 函数在规定时间执行完后关闭alarm闹钟
            return result
        
        if platform.system() == "Windows":
            return func

        return wrapper

    return decorator