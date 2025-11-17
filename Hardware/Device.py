
# ==================== Device.py ====================
# File heads are executed when `import Device' is called.
# the behavior of 'import' ensures logger only initialized once.
# logger.handler enqueue ensures log file is not locked by other process.


# Reconfigure logger
import os, sys, warnings, inspect

# ------------ Logger start ------------
from loguru import logger

def get_call_kwargs(level=1):
    frame = inspect.currentframe().f_back
    for _ in range(level):  # get the level-th frame
        frame = frame.f_back
    code_obj = frame.f_code
    return dict(
        function_module=os.path.basename(code_obj.co_filename),
        function_name=code_obj.co_name,
        function_line=frame.f_lineno,
    )

def send_log_file_via_email(fname):
    import win32com.client
    ol = win32com.client.Dispatch('Outlook.Application')
    # size of the new email
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = '[Regular] [Keck LFC] log file rotated: ' + str(os.path.split(fname)[1])
    newmail.To = 'maodonggao@outlook.com; stephanie.leifer@aero.org; jge2@caltech.edu'
    # newmail.CC='maodonggao@outlook.com'
    newmail.Body = '[Automatically Generated Email] \n\n' + 'Hello, \n\n Attached are the rotated data logging file at Keck. Email sent for backup only. \n\n Best, \n Maodong'

    newmail.Attachments.Add(fname)
    newmail.Send()
    print(f'Log file {fname} sent')


# 移除默认的 logger
logger.remove()

# 日志输出格式
logger_format = (
    "<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green> | "
    "<level>{level: <8}</level> | "
    "<cyan>{extra[devicename]}</cyan> | "
    "<cyan>{extra[function_module]}</cyan>:"
    "<cyan>{extra[function_name]}</cyan>:"
    "<cyan>{extra[function_line]}</cyan>\n"
    "<level>{message}</level>"
)
log_template = os.path.expanduser(
    r'~\Desktop\Keck\Logs\test_{time:YYYY-MM-DD_HH-mm-ss}.log'
)

logger.add(sys.stderr, format=logger_format, level="INFO")  # recover console print
logger.add(log_template, format=logger_format, level="INFO", rotation="1 MB", retention=50, enqueue=True,
           compression=send_log_file_via_email)  # 1MB per file, 50 files max
logger.bind(devicename="Device").info('logger initialized', **get_call_kwargs(level=0))

# ------------ Logger end ------------


#Base class for devices
# class Device:
#     import pyvisa
#     rm = pyvisa.ResourceManager()

#     def __init__(self, addr, name='', isVISA=True):
#         self.addr = addr
#         self.devicename = name
#         self.isVISA = isVISA
#         self.connected = False
#         if isVISA:
#             try:
#                 self.inst = self.rm.open_resource(addr)
#             except:  # TODO： raise warning (or error?) here to help identify visa object create failure.
#                 pass

#     def connect(self):
#         if not self.connected:
#             try:
#                 if self.isVISA:
#                     self.inst.open()
#                 self.connected = True
#                 self.info(self.devicename + " connected")
#                 return 1
#             except Exception as e:
#                 self.error(f"Error:{e}")
#                 return -1
#         return 0

#     def disconnect(self):
#         if self.connected:
#             if self.isVISA:
#                 self.inst.close()
#             self.connected = False
#             self.info(self.devicename + " disconnected")
#             return 1
#         return 0

#     def write(self, cmd):
#         self.inst.write(cmd)

#     def read(self):
#         return self.inst.read()

#     def query(self, cmd):
#         return self.inst.query(cmd)

# class Device:
#     import pyvisa
#     import time
#     import warnings
#     from loguru import logger
#     #from .utils import get_call_kwargs
#     rm = pyvisa.ResourceManager()

#     def __init__(self, addr, name='', isVISA=True):
#         self.addr = addr
#         self.devicename = name
#         self.isVISA = isVISA
#         self.connected = False
#         self.inst = None
#         if isVISA:
#             self._open_visa()

#     def _open_visa(self):
#         try:
#             if self.inst is not None:
#                 try:
#                     self.inst.close()
#                 except Exception:
#                     pass
#             self.inst = self.rm.open_resource(self.addr)
#             self.connected = True
#             self.info(self.devicename + " connected")
#         except Exception as e:
#             self.connected = False
#             self.inst = None
#             self.error(f"Error opening VISA resource: {e}")

#     def connect(self):
#         if not self.connected:
#             self._open_visa()
#         return 1 if self.connected else -1

#     def disconnect(self):
#         if self.connected and self.inst is not None:
#             try:
#                 self.inst.close()
#             except Exception:
#                 pass
#             self.connected = False
#             self.info(self.devicename + " disconnected")
#             return 1
#         return 0

#     def __del__(self):
#         self.disconnect()

#     def _auto_reconnect(func):
#         """装饰器：通信失败时自动重连一次"""
#         def wrapper(self, *args, **kwargs):
#             if not self.connected:
#                 self.connect()
#             try:
#                 return func(self, *args, **kwargs)
#             except (pyvisa.errors.VisaIOError, AttributeError) as e:
#                 self.warning(f"VISA IO error or not connected: {e}, attempting reconnect...")
#                 self.disconnect()
#                 time.sleep(1)
#                 self.connect()
#                 try:
#                     return func(self, *args, **kwargs)
#                 except Exception as e2:
#                     self.error(f"Failed after reconnect: {e2}")
#                     raise e2
#         return wrapper

#     @_auto_reconnect
#     def write(self, cmd):
#         return self.inst.write(cmd)

#     @_auto_reconnect
#     def read(self):
#         return self.inst.read()

#     @_auto_reconnect
#     def query(self, cmd):
#         return self.inst.query(cmd)

class Device:
    import pyvisa
    rm = pyvisa.ResourceManager()

    def __init__(self, addr, name='', isVISA=True):
        self.addr = addr
        self.devicename = name
        self.isVISA = isVISA
        self.connected = False
        self.inst = None  # 确保属性存在，兼容子类在 __init__ 里直接用 self.inst.xxx
        if isVISA:
            try:
                self.inst = self.rm.open_resource(addr)  # 保持原始行为
            except Exception as e:
                # 保持静默（兼容旧逻辑），但至少不致使 self.inst 缺失
                # 如需调试可临时开启：self.warning(f"VISA open_resource failed in __init__: {e}")
                self.inst = None

    def connect(self):
        if not self.connected:
            try:
                if self.isVISA:
                    # 兼容原逻辑：__init__ 已 open_resource；connect 再 open()
                    if self.inst is None:
                        # 若 __init__ 失败，这里补一次 open_resource（不改变旧成功路径）
                        self.inst = self.rm.open_resource(self.addr)
                    # 某些资源类型没有 open()，容错处理
                    try:
                        self.inst.open()
                    except Exception:
                        pass
                # 非 VISA：保持原行为——只标记 connected，不做任何 pyvisa 操作
                self.connected = True
                self.info(self.devicename + " connected")
                return 1
            except Exception as e:
                self.error(f"Error:{e}")
                return -1
        return 0

    def disconnect(self):
        if self.connected:
            if self.isVISA and self.inst is not None:
                try:
                    self.inst.close()
                except Exception:
                    pass
            self.connected = False
            self.info(self.devicename + " disconnected")
            return 1
        return 0

    # ---- 内部工具：复用已设 I/O 参数（超时/终止符/波特率等）以保持子类原有设置 ----
    def _snapshot_io_attrs(self):
        attrs = {}
        inst = getattr(self, "inst", None)
        if inst is None:
            return attrs
        keys = ("timeout", "read_termination", "write_termination",
                "baud_rate", "chunk_size", "encoding")
        for k in keys:
            if hasattr(inst, k):
                try:
                    attrs[k] = getattr(inst, k)
                except Exception:
                    pass
        return attrs

    def _restore_io_attrs(self, attrs):
        inst = getattr(self, "inst", None)
        if inst is None:
            return
        for k, v in attrs.items():
            try:
                setattr(inst, k, v)
            except Exception:
                pass

    def _visa_io_with_reconnect(self, op_callable):
        """
        仅对 VISA 设备：执行 I/O，失败一次后自动重连并重试一次。
        - 成功则与过去完全一致；
        - 失败时才触发兜底；
        - 重连后恢复旧的 I/O 参数，并调用子类 `_after_reconnect()`（若存在）。
        """
        import time as _time
        if not self.isVISA:
            # 非 VISA 不做任何包装，保持旧逻辑
            return op_callable()

        # 若尚未 connect，也按旧逻辑补一次
        if not self.connected or self.inst is None:
            self.connect()

        try:
            return op_callable()

        except Exception as e1:
            # 仅在异常时重连一次；尽量保留之前的 I/O 参数
            io_attrs = self._snapshot_io_attrs()
            self.warning(f"{self.devicename} I/O failed: {e1}; reconnecting once...")

            try:
                # 关闭旧会话
                try:
                    if self.inst is not None:
                        self.inst.close()
                except Exception:
                    pass

                # 重新打开
                self.inst = self.rm.open_resource(self.addr)
                try:
                    self.inst.open()
                except Exception:
                    pass
                self.connected = True

                # 恢复 I/O 参数（超时/终止符/波特率等）
                self._restore_io_attrs(io_attrs)

                # >>> 新增：允许设备自行做串口后处理（清缓冲/再设串口参数/小延时等）
                if hasattr(self, "_after_reconnect"):
                    try:
                        self._after_reconnect()
                    except Exception as _e:
                        self.warning(f"_after_reconnect failed: {_e}")

                _time.sleep(0.2)
                return op_callable()

            except Exception as e2:
                self.error(f"{self.devicename} I/O failed after reconnect: {e2}")
                raise

    # ---- 公共 I/O 接口：保持旧签名；仅在 VISA 时加“自动重连”兜底 ----
    def write(self, cmd):
        if self.isVISA:
            return self._visa_io_with_reconnect(lambda: self.inst.write(cmd))
        # 非 VISA：保持原逻辑直接调用
        return self.inst.write(cmd)

    def read(self):
        if self.isVISA:
            return self._visa_io_with_reconnect(lambda: self.inst.read())
        return self.inst.read()

    def query(self, cmd):
        if self.isVISA:
            return self._visa_io_with_reconnect(lambda: self.inst.query(cmd))
        return self.inst.query(cmd)

    # logging
    logger = logger
    
    def debug(self, x, name='', level=1):
        logger.debug(x, devicename=name or self.devicename, **get_call_kwargs(level))
        # print(x)

    def info(self, x, name='', level=1):
        logger.info(x, devicename=name or self.devicename, **get_call_kwargs(level))
        # print(x)

    def warning(self, x, name='', level=1):
        logger.warning(x, devicename=name or self.devicename, **get_call_kwargs(level))
        warnings.warn(x)

    def error(self, x, name='', level=1):
        logger.error(x, devicename=name or self.devicename, **get_call_kwargs(level))
        # raise Exception(x)
# ------------ Device base class (VISA + non-VISA) ------------



###########3333
# import os
# import sys
# import warnings
# import inspect
# from loguru import logger
# import time

# def get_call_kwargs(level=1):
#     frame = inspect.currentframe().f_back
#     for _ in range(level):  # get the level-th frame
#         frame = frame.f_back
#     code_obj = frame.f_code
#     return dict(
#         function_module=os.path.basename(code_obj.co_filename),
#         function_name=code_obj.co_name,
#         function_line=frame.f_lineno,
#     )

# # Custom function to handle log file sending and retry log rotation
# def send_log_file_via_email(fname):
#     import win32com.client
#     ol = win32com.client.Dispatch('Outlook.Application')
#     olmailitem = 0x0
#     newmail = ol.CreateItem(olmailitem)
#     newmail.Subject = '[Regular] [Keck LFC] log file rotated: ' + str(os.path.split(fname)[1])
#     newmail.To = 'maodonggao@outlook.com; stephanie.leifer@aero.org; jge2@caltech.edu'
#     newmail.Body = '[Automatically Generated Email] \n\n' + 'Hello, \n\n Attached are the rotated data logging file at Keck. Email sent for backup only. \n\n Best, \n Maodong'
#     newmail.Attachments.Add(fname)
#     newmail.Send()
#     print(f'Log file {fname} sent')

# # Retry function for rotating the log if locked
# def safe_rotate_log(file_path, retries=3, wait_time=5):
#     for attempt in range(retries):
#         try:
#             os.rename(file_path, f"{file_path}.{time.strftime('%Y-%m-%d_%H-%M-%S')}")
#             send_log_file_via_email(file_path)
#             break
#         except PermissionError:
#             print(f"Permission error encountered during log rotation. Retrying... ({attempt + 1}/{retries})")
#             time.sleep(wait_time)
#     else:
#         print(f"Failed to rotate log after {retries} attempts.")

# # Logger reconfiguration
# fname = os.path.expanduser(r'~\Desktop\Keck\Logs\test.log')
# logger.remove()  # Remove default logger
# logger_format = (
#     "<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green> | "
#     "<level>{level: <8}</level> | "
#     "<cyan>{extra[devicename]}</cyan> | "
#     "<cyan>{extra[function_module]}</cyan>:<cyan>{extra[function_name]}</cyan>:<cyan>{extra[function_line]}</cyan>\n"
#     "<level>{message}</level>")

# logger.add(sys.stderr, format=logger_format, level="INFO")  # Console logger
# logger.add(
#     fname, format=logger_format, level="INFO", 
#     rotation="1 MB", retention=5, enqueue=True, 
#     compression=lambda f: safe_rotate_log(f)  # Use safe log rotation function
# )
# logger.bind(devicename="Device").info('logger initialized', **get_call_kwargs(level=0))

# Now any logging operation will safely rotate logs and handle errors.


# import psutil

# def kill_process_using_file(file_path):
#     for proc in psutil.process_iter(['pid', 'name']):
#         try:
#             for file in proc.open_files():
#                 if file.path == file_path:
#                     print(f"Process {proc.info['name']} (PID {proc.info['pid']}) is using {file_path}")
#                     proc.kill()  # 杀死占用文件的进程
#         except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#             continue

# kill_process_using_file(r'C:\Users\KeckLFC\Desktop\Keck\Logs\test.log')
