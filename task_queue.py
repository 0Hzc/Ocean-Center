#!/usr/bin/env python3
"""
任务队列管理器
用于防止多个实例同时运行,实现排队机制
"""
import os
import fcntl
import time
from datetime import datetime


class TaskQueue:
    """
    基于文件锁的任务队列管理器
    确保同一时间只有一个任务在运行
    """
    def __init__(self, lock_file='/tmp/ocean_center.lock'):
        """
        初始化任务队列

        :param lock_file: 锁文件路径
        """
        self.lock_file = lock_file
        self.lock_fd = None
        self.acquired = False

    def acquire(self, timeout=300, wait=True):
        """
        获取锁,如果获取失败则等待

        :param timeout: 超时时间(秒),默认300秒
        :param wait: 是否等待锁释放,如果False则立即返回
        :return: 是否成功获取锁
        """
        try:
            # 确保锁文件所在目录存在
            lock_dir = os.path.dirname(self.lock_file)
            if lock_dir and not os.path.exists(lock_dir):
                os.makedirs(lock_dir, exist_ok=True)

            # 打开或创建锁文件
            self.lock_fd = open(self.lock_file, 'w')

            # 尝试非阻塞方式获取锁
            try:
                fcntl.flock(self.lock_fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
                self.acquired = True
                # 写入当前进程信息
                self.lock_fd.write(f"PID: {os.getpid()}\n")
                self.lock_fd.write(f"时间: {datetime.now().isoformat()}\n")
                self.lock_fd.flush()
                print(f"[TaskQueue] 成功获取任务锁")
                return True
            except IOError:
                # 锁已被其他进程持有
                if not wait:
                    self.lock_fd.close()
                    self.lock_fd = None
                    print(f"[TaskQueue] 任务正在运行中,无法获取锁")
                    return False

                # 等待锁释放
                print(f"[TaskQueue] 检测到其他任务正在运行,进入等待队列...")
                print(f"[TaskQueue] 最长等待时间: {timeout}秒")
                start_time = time.time()

                # 阻塞方式等待锁
                fcntl.flock(self.lock_fd, fcntl.LOCK_EX)

                wait_time = time.time() - start_time
                print(f"[TaskQueue] 等待了 {wait_time:.1f} 秒后成功获取任务锁")

                self.acquired = True
                # 写入当前进程信息
                self.lock_fd.write(f"PID: {os.getpid()}\n")
                self.lock_fd.write(f"时间: {datetime.now().isoformat()}\n")
                self.lock_fd.write(f"等待时间: {wait_time:.1f}秒\n")
                self.lock_fd.flush()
                return True

        except Exception as e:
            print(f"[TaskQueue] 获取锁时发生错误: {e}")
            if self.lock_fd:
                self.lock_fd.close()
                self.lock_fd = None
            return False

    def release(self):
        """释放锁"""
        if self.acquired and self.lock_fd:
            try:
                fcntl.flock(self.lock_fd, fcntl.LOCK_UN)
                self.lock_fd.close()
                self.lock_fd = None
                self.acquired = False
                print(f"[TaskQueue] 任务锁已释放")
            except Exception as e:
                print(f"[TaskQueue] 释放锁时发生错误: {e}")

    def __enter__(self):
        """支持with语句"""
        if self.acquire():
            return self
        else:
            raise RuntimeError("无法获取任务锁")

    def __exit__(self, exc_type, exc_val, exc_tb):
        """支持with语句"""
        self.release()
        return False


if __name__ == "__main__":
    # 测试代码
    print("测试任务队列...")

    with TaskQueue() as queue:
        print("获取到锁,模拟任务执行...")
        time.sleep(5)
        print("任务执行完成")

    print("测试完成")
