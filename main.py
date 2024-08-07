from shutil import copyfile
from typing import List

import pandas as pd
import os
import time
import tkinter as tk
from tkinter import messagebox
import threading


def generate_file_path(file_name, channel_num):
    file_path_list = []
    date_list = pd.date_range(start_date, end_date, freq="D").format(
        date_format="%Y%m%d"
    )
    for project_name in project_name_list:
        for mac_addr in mac_addr_list:
            mac_dir = os.path.join(remote_url, project_name, mac_addr)
            if not os.path.exists(mac_dir):
                continue
            real_date_dir_list = os.listdir(mac_dir)
            for date in date_list:
                date_dir = os.path.join(mac_dir, date)
                if date not in real_date_dir_list:
                    continue
                for channel in range(channel_num):
                    file_path = (
                        f"{remote_url}\\{project_name}\\{mac_addr}\\{date}\\"
                        f"{date}_{channel:02d}_E_{file_name}.csv"
                    )
                    file_path_list.append(file_path)
    return file_path_list


def translate_path(src_path):
    words = src_path.split("\\")
    file_name = words[-1].split(".")[0]
    words2 = file_name.split("_E_")
    mac_addr = words[-3]
    project_name = words[-4]
    new_path = os.path.join(
        data_dir, f"{words2[1]}_{words2[0]}_{project_name}_{mac_addr}.csv"
    )
    return new_path


def download_csv():
    lb_filename_on_focus_out()
    if len(file_name_list) == 0:
        b_download.config(state="active")
        return
    if not os.path.exists(data_dir):
        os.mkdir(data_dir)
    channel_num = int(channel_num_var.get())
    file_path_list_list: List[List[str]] = []
    count = 0
    for file_name in file_name_list:
        file_path_list: List[str] = generate_file_path(file_name, channel_num)
        count += len(file_path_list)
        file_path_list_list.append(file_path_list)

    current_count = 0
    i = 0
    for file_path_list in file_path_list_list:
        if not option_merge:
            for file_path in file_path_list:
                current_count += 1
                str_progress.set(f"{current_count}/{count}")
                if os.path.exists(file_path):
                    copyfile(file_path, translate_path(file_path))
        elif option_merge:
            all_data_frame = []
            for file_path in file_path_list:
                current_count += 1
                str_progress.set(f"{current_count}/{count}")
                if os.path.exists(file_path):
                    try:
                        data_frame = pd.read_csv(file_path)
                    except Exception:
                        with open(data_dir + "\\error.txt", "a") as f:
                            f.write(file_path + "\n")
                        i += 1
                        continue
                    else:
                        all_data_frame.append(data_frame)
            data_frame_concat = pd.concat(all_data_frame, axis=0, ignore_index=True)
            data_frame_concat.to_csv(os.path.join(data_dir, file_name_list[i] + ".csv"))
        i += 1
    messagebox.showinfo(message="下载完成")
    b_download.config(state="active")
    os.system(f"explorer.exe {data_dir}")


def async_download_csv():
    b_download.config(state="disabled")
    th = threading.Thread(target=download_csv)
    th.start()


def search_project():
    keyword = e_project_key.get()
    lb_project_list.delete(0, "end")
    for project_name in all_project_name:
        if project_name.find(keyword) != -1:
            lb_project_list.insert("end", project_name)


def step1():
    project_name_list.clear()
    for i in lb_project_list.curselection():
        project_name_list.append(lb_project_list.get(i))
    if len(project_name_list) != 0:
        mac_list = []
        for project_name in project_name_list:
            mac_addr_t = os.listdir(os.path.join(remote_url, project_name))
            mac_list.extend(mac_addr_t)
        mac_set_list = list(set(mac_list))
        lb_mac_addr.delete(0, "end")
        for mac in mac_set_list:
            lb_mac_addr.insert("end", mac)

        b_step2.config(state="active")


def step2():
    mac_addr_list.clear()
    all_date_list = []
    for i in lb_mac_addr.curselection():
        mac_addr_list.append(lb_mac_addr.get(i))
    if len(mac_addr_list) != 0:
        date_dir = os.path.join(remote_url, project_name_list[0], mac_addr_list[0])
        for project_name in project_name_list:
            for mac_addr in mac_addr_list:
                date_dir = os.path.join(remote_url, project_name, mac_addr)
                if os.path.exists(date_dir):
                    date_list = os.listdir(date_dir)
                    all_date_list.extend(date_list)
        sorted_date_list = sorted(set(all_date_list))
        e_date_start.delete(0, "end")
        e_date_start.insert(0, sorted_date_list[0])
        e_date_end.delete(0, "end")
        e_date_end.insert(0, sorted_date_list[-1])
        b_step3.config(state="active")


def split_name(fullname):
    return fullname.split(".")[0].split("_E_")[-1]


def get_max_channel_num(file_list):
    max_channel = 0
    for filename in file_list:
        channel_num = int(filename.split("_")[1])
        if channel_num > max_channel:
            max_channel = channel_num

    return max_channel + 1


def step3():
    lb_filename_list.delete(0, "end")
    global start_date
    global end_date
    start_date = e_date_start.get()
    end_date = e_date_end.get()
    file_path = os.path.join(
        remote_url, project_name_list[0], mac_addr_list[0], start_date
    )
    e_file_url.delete(0, "end")
    e_file_url.insert(0, file_path)
    if os.path.exists(file_path):
        file_list = os.listdir(file_path)
        max_channel = get_max_channel_num(file_list)
        channel_num_var.set(max_channel)
        short_filename = []
        for filename in file_list:
            short_filename.append(split_name(filename))
        short_filename = list(set(short_filename))
        for filename in short_filename:
            lb_filename_list.insert("end", filename)
    b_download.config(state="active")
    b_refresh.config(state="active")


def refresh():
    file_path = e_file_url.get()
    lb_filename_list.delete(0, "end")
    if os.path.exists(file_path):
        file_list = os.listdir(file_path)
        max_channel = get_max_channel_num(file_list)
        channel_num_var.set(max_channel)
        short_filename = []
        for filename in file_list:
            short_filename.append(split_name(filename))
        short_filename = list(set(short_filename))
        for filename in short_filename:
            lb_filename_list.insert("end", filename)


def on_vertical(event=None):
    print(s.get())  # 获取当前滑块的位置
    print(event.delta)  # 向上还是向下滑动
    if event.delta > 0:
        c.yview_moveto(round(s.get()[0] - 0.1, 1))
    elif event.delta < 0:
        c.yview_moveto(round(s.get()[0] + 0.1, 1))


def lb_filename_on_focus_out(event=None):
    if len(lb_filename_list.curselection()) == 0:
        return
    file_name_list.clear()
    for i in lb_filename_list.curselection():
        file_name_list.append(lb_filename_list.get(i))


if __name__ == "__main__":
    remote_url = "\\\\192.168.12.245\\csv"
    project_name_list = []
    mac_addr_list = []
    start_date = ""
    end_date = ""
    file_name_list = []
    # CSV合并选项
    option_merge = True
    data_dir = time.strftime("%Y%m%d_%H%M%S")

    all_project_name = os.listdir(remote_url)

    # 第1步，实例化object，建立窗口window
    root = tk.Tk()

    s = tk.Scrollbar(root)
    s.pack(side="right", fill="y")

    # 第2步，给窗口的可视化起名字
    root.title("CSV Download v2.2.0")

    root.geometry("800x960")

    c = tk.Canvas(root)
    c.pack(fill="both", expand=1)

    window = tk.Frame(c, width=800, height=1100)

    c.create_window(10, 800, window=window, anchor="nw")
    c.update()

    c.config(yscrollcommand=s.set, scrollregion=c.bbox("all"))
    s.config(command=c.yview)
    c.bind("<MouseWheel>", on_vertical)  # 框架绑定鼠标滑动事件
    window.bind("<MouseWheel>", on_vertical)  # 框架绑定鼠标滑动事件

    project_frame = tk.Frame(window)

    project_frame.pack()

    e_project_key = tk.Entry(project_frame)
    e_project_key.pack(side="left")

    b_search_project = tk.Button(project_frame, text="搜索项目", command=search_project)
    b_search_project.pack(side="left")

    l_project_list = tk.Label(window, text="项目列表")
    l_project_list.pack()

    lb_project_list = tk.Listbox(window, selectmode="extended", width=80)
    lb_project_list.pack()

    b_step1 = tk.Button(window, text="Step1", command=step1)
    b_step1.pack()

    l_mac_addr = tk.Label(window, text="MAC地址")
    l_mac_addr.pack()

    lb_mac_addr = tk.Listbox(window, selectmode="extended", width=80)
    lb_mac_addr.pack()

    b_step2 = tk.Button(window, text="Step2", command=step2)
    b_step2.pack()
    b_step2.config(state="disabled")

    l_date = tk.Label(window, text="起止日期")
    l_date.pack()

    date_frame = tk.Frame(window)
    date_frame.pack()

    e_date_start = tk.Entry(date_frame)
    e_date_end = tk.Entry(date_frame)

    l_split = tk.Label(date_frame, text="~")

    e_date_start.pack(side="left")
    l_split.pack(side="left")
    e_date_end.pack(side="left")

    b_step3 = tk.Button(window, text="Step3", command=step3)
    b_step3.pack()
    b_step3.config(state="disabled")

    l_filename_list = tk.Label(window, text="文件列表")
    l_filename_list.pack()

    lb_filename_list = tk.Listbox(window, selectmode="extended", width=80)
    lb_filename_list.pack()
    lb_filename_list.bind("<FocusOut>", lb_filename_on_focus_out)

    refresh_frame = tk.Frame(window)
    refresh_frame.pack()
    e_file_url = tk.Entry(refresh_frame)
    e_file_url.pack(side="left")
    b_refresh = tk.Button(refresh_frame, text="从url刷新", command=refresh)
    b_refresh.pack()

    channel_num_frame = tk.Frame(window)
    channel_num_frame.pack()
    l_channel_num = tk.Label(channel_num_frame, text="通道数量")
    l_channel_num.pack(side="left")
    channel_num_var = tk.StringVar(value="4")
    e_channel_num = tk.Entry(channel_num_frame, textvariable=channel_num_var)
    e_channel_num.pack(side="left")

    b_download = tk.Button(window, text="下载", command=async_download_csv)
    b_download.pack()
    b_download.config(state="disabled")

    str_progress = tk.StringVar()
    str_progress.set("0/0")
    l_progress = tk.Label(window, textvariable=str_progress)
    l_progress.pack()

    # 第6步，主窗口循环显示
    root.mainloop()
