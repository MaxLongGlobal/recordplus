import win32api
import win32con
from pynput.keyboard import Listener, Key
import pyaudio
import wave
from PIL import ImageGrab
import cv2
import threading
import time
from numpy import array
from moviepy.editor import *
import os
from win32api import MessageBox

flag = False  # 停止标志位


class PyRecord:
    def __init__(self):
        self.allow_record = True
        path = make_file_dir()
        self.file_path = path + r"/{}".format(time.strftime('%Y-%m-%d %H-%M-%S'))

    def record_audio(self):
        # 如无法正常录音 请启用计算机的"立体声混音"输入设备
        CHUNK = 1024
        FORMAT = pyaudio.paInt16
        CHANNELS = 2
        RATE = 11025
        p = pyaudio.PyAudio()
        stream = p.open(
            format=FORMAT,
            channels=CHANNELS,
            rate=RATE,
            input=True,
            frames_per_buffer=CHUNK,
        )
        wf = wave.open(self.file_path + ".wav", "wb")
        wf.setnchannels(CHANNELS)
        wf.setsampwidth(p.get_sample_size(FORMAT))
        wf.setframerate(RATE)

        while self.allow_record:
            if flag:
                print("record_audio 录制结束！")
                # MessageBox(0, "        退出录制 点击确定", "全屏录制工具")
                break
            data = stream.read(CHUNK)
            wf.writeframes(data)

        stream.stop_stream()
        stream.close()
        p.terminate()
        wf.close()

    def record_screen(self):
        im = ImageGrab.grab()
        video = cv2.VideoWriter(
            self.file_path + ".mp4", cv2.VideoWriter_fourcc(*"XVID"), 10, im.size
        )
        while self.allow_record:
            if flag:
                print("allow_record 录制结束！")
                # MessageBox(0, "        退出录制 点击确定", "全屏录制工具")
                break
            im = ImageGrab.grab()
            im = cv2.cvtColor(array(im), cv2.COLOR_RGB2BGR)
            video.write(im)
        video.release()

    def compose_file(self):
        print("合并视频&音频文件")
        audio = AudioFileClip(self.file_path + ".wav")
        video = VideoFileClip(self.file_path + ".mp4")
        ratio = audio.duration / video.duration
        video = video.fl_time(lambda t: t / ratio, apply_to=["video"]).set_end(
            audio.duration
        )
        video = video.set_audio(audio)
        video = video.volumex(5)
        video.write_videofile(self.file_path + "_out.mp4".format(time.strftime('%Y-%m-%d %H-%M-%S')), codec="libx264", fps=25, logger=None)
        video.close()

    def remove_temp_file(self):
        print("删除缓存文件")
        os.remove(self.file_path + ".wav")
        os.remove(self.file_path + ".mp4")

    def stop(self):
        print("停止录制")
        self.allow_record = False
        time.sleep(1)
        self.compose_file()
        self.remove_temp_file()

    def run(self):
        t = threading.Thread(target=self.record_screen)
        t1 = threading.Thread(target=self.record_audio)
        t.start()
        t1.start()
        print("开始录制")
        with Listener(on_press=on_press) as listener:
            listener.join()


def on_press(key):
    """
    键盘监听事件！！！
    :param key:
    :return:
    """
    global flag
    if key == Key.esc:
        flag = True
        print("stop monitor！")
        return False  # 返回False，键盘监听结束！


def Minimize_Window():
    win32api.keybd_event(91, 0, 0, 0)
    time.sleep(0.5)
    win32api.keybd_event(40, 0, 0, 0)
    time.sleep(0.5)
    win32api.keybd_event(91, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(40, 0, win32con.KEYEVENTF_KEYUP, 0)


def make_file_dir():
    path = r"video"
    is_exist = os.path.exists(path)
    # 判断结果
    if not is_exist:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
        print('创建目录')
        return path
    else:
        # 如果目录存在则不创建，并提示目录已存在
        print('目录已存在')
        return path


if __name__ == '__main__':
    # Minimize_Window()
    a = MessageBox(0, "        开始录制 点击确定\n        结束录制 按下Esc", "全屏录制工具")
    if a == 1:
        time.sleep(2)
        pr = PyRecord()
        pr.run()
        pr.stop()
    b = MessageBox(0, "        退出录制 点击确定", "全屏录制工具")
    if b == 1:
        print('退出录制')
