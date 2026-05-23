import tkinter as tk
from tkinter import messagebox
from tkinter import simpledialog
from tkinter import ttk
from tkinter import filedialog
import json
import os
import time
import random
import threading
#from rich.console import Console

class Console:
    def ANSIcolor(self, color="#000000"):
        if color.startswith("#") and len(color) == 7:
            r = int(color[1:3], 16)
            g = int(color[3:5], 16)
            b = int(color[5:7], 16)
            return f"\033[38;2;{r};{g};{b}m"
        return "\033[0m"

    def print(self, *args, sep=" ", end="\n", style="#000000"):
        args = [str(arg) for arg in args]
        text = sep.join(args)
        ANSIcolor = self.ANSIcolor(style)
        ANSICode = f"{ANSIcolor}{text}{end}\033[0m"
        print(ANSICode, end="")

CONSOLE = Console()

def timer():
    def wrapper(func):
        def inner(*args, **kwargs):
            start_time = time.time()
            result = func(*args, **kwargs)
            end_time = time.time()
            CONSOLE.print(f"[*] {func.__name__} executed in {end_time - start_time:.4f} seconds", style="#00bb00")
            return result
        return inner
    return wrapper

class Explain:
    def __init__(self, parent, text, control, font="assets/MicrosoftYaHei.ttc", size=12, padx=0, pady=4):
        self.parent = parent
        self.text = text
        self.control = control
        self.font = font
        self.size = size
        self.padx = padx
        self.pady = pady

        self.label = tk.Label(
            parent,
            text=text,
            font=(font, size),
            fg="#555"
        )
        self._place_label()

    def _place_label(self):
        self.parent.update_idletasks()
        if not self.control.winfo_ismapped():
            self.parent.after(10, self._place_label)
            return

        x = self.control.winfo_x()
        y = self.control.winfo_y()
        self.label.place(x=x + self.padx, y=y - self.label.winfo_reqheight() - self.pady)

class MainUI():
    def __init__(self):
        self.showInfo()
#------------------Init Variable-----------------------#
        self.latestPos = {"x": 0, "y": 0}
        self.font = "assets/MicrosoftYaHei.ttc"
        self.choiseList = []
        self.maxcount = 1
        self.current_file_path = None
        self.countVar = None
        self.roll_job = None
        self.rolling = False
        self.roll_target_count = 1
        self.choice_folder = "assets/choicefile"
        os.makedirs(self.choice_folder, exist_ok=True)
#---------------------Init UI--------------------------#
        self.root = tk.Tk()
        self.countVar = tk.IntVar(self.root, value=1)
        self.root.title("choice UI")
        self.root.resizable(False, False)
        
        self.Controls()
        self.LoadSetting()
        self.root.iconbitmap("assets/icon.ico")
        self.threadPrograss(self.UIposChanged)

        self.root.mainloop()

    def Controls(self):
        self.showLabel = tk.Label(self.root, text="等待抽取", font=(self.font, 20))
        self.showLabel.place(x=400, y=10, anchor="n")

        self.startButton = ttk.Button(self.root, text="开始抽奖", command=self.Choice)
        self.startButton.place(x=300, y=60, anchor="n")

        self.setButton = ttk.Button(self.root, text="设置选项", command=self.Setting)
        self.setButton.place(x=500, y=60, anchor="n")

        self.countSpinBox = ttk.Spinbox(self.root, from_=1, to=100, textvariable=self.countVar, width=5)
        self.countSpinBox.place(x=300, y=150, anchor="n", width=100)
        Explain(self.root, "抽取个数", self.countSpinBox)

        self.countInfoLabel = tk.Label(self.root, text="名单数量：0", font=(self.font, 10), fg="#333")
        self.countInfoLabel.place(x=420, y=150, anchor="n")

        self.loadedFileLabel = tk.Label(self.root, text="当前名单：无", font=(self.font, 10), fg="#333")
        self.loadedFileLabel.place(x=420, y=180, anchor="n")

        self.UpdateCountLimit()

    def showInfo(self):

        def show():
            with open("assets/CONTENT.txt", "r", encoding="utf-8") as f:
                info = f.read()
            CONSOLE.print(info, style="#ffff00")

        show()

    def Setting(self):
        class SettingUI():
            def __init__(self, parent):
                self.parent = parent
                self.setroot = tk.Toplevel(parent.root)
                self.setroot.title("设置选项")
                self.setroot.resizable(False, False)
                self.setroot.iconbitmap("assets/icon.ico")
                self.setroot.geometry("520x360")
                self.Setcontrols()
                self.RefreshList()
                self.setroot.grab_set()

            def Setcontrols(self):
                self.importButton = ttk.Button(self.setroot, text="导入TXT名单", command=self.parent.askFile)
                self.importButton.place(x=80, y=20, anchor="n", width=120)

                self.newButton = ttk.Button(self.setroot, text="新建空名单", command=self.CreateEmptyList)
                self.newButton.place(x=200, y=20, anchor="n", width=120)

                self.refreshButton = ttk.Button(self.setroot, text="刷新列表", command=self.RefreshList)
                self.refreshButton.place(x=320, y=20, anchor="n", width=120)

                self.renameButton = ttk.Button(self.setroot, text="重命名名单", command=self.RenameSelectedFile)
                self.renameButton.place(x=440, y=20, anchor="n", width=120)

                self.fileListbox = tk.Listbox(self.setroot, height=12, width=45)
                self.fileListbox.place(x=260, y=50, anchor="n")

                self.loadButton = ttk.Button(self.setroot, text="加载名单", command=self.LoadSelectedFile)
                self.loadButton.place(x=80, y=300, anchor="n", width=120)

                self.editButton = ttk.Button(self.setroot, text="编辑名单", command=self.OpenEditor)
                self.editButton.place(x=200, y=300, anchor="n", width=120)

                self.exportButton = ttk.Button(self.setroot, text="导出TXT", command=self.ExportSelectedFile)
                self.exportButton.place(x=320, y=300, anchor="n", width=120)

                self.deleteButton = ttk.Button(self.setroot, text="删除名单", command=self.DeleteSelectedFile)
                self.deleteButton.place(x=440, y=300, anchor="n", width=120)

            def RefreshList(self):
                self.fileListbox.delete(0, tk.END)
                files = self.parent.LoadChoiceFiles()
                for f in files:
                    self.fileListbox.insert(tk.END, f)

            def GetSelectedFile(self):
                selection = self.fileListbox.curselection()
                if not selection:
                    CONSOLE.print("[-] No file selected.", style="#bbbb00")
                    messagebox.showwarning("提示", "请先选择一个名单文件。")
                    return None
                return self.fileListbox.get(selection[0])

            def LoadSelectedFile(self):
                filename = self.GetSelectedFile()
                if not filename:
                    return
                file_path = os.path.join(self.parent.choice_folder, filename)
                self.parent.LoadFile(file_path)
                CONSOLE.print(f"[*] Loaded file: {filename}", style="#00bb00")
                messagebox.showinfo("加载成功", f"已加载名单：{filename}")

            def OpenEditor(self):
                filename = self.GetSelectedFile()
                if not filename:
                    return
                file_path = os.path.join(self.parent.choice_folder, filename)
                FileEditor(self.parent, file_path)

            def ExportSelectedFile(self):
                filename = self.GetSelectedFile()
                if not filename:
                    return
                file_path = os.path.join(self.parent.choice_folder, filename)
                save_path = filedialog.asksaveasfilename(
                    title="导出TXT",
                    defaultextension=".txt",
                    filetypes=[("Text files", "*.txt")],
                    initialfile=os.path.splitext(filename)[0] + ".txt"
                )
                if not save_path:
                    return
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        names = json.load(f)
                    with open(save_path, "w", encoding="utf-8") as f:
                        for name in names:
                            f.write(str(name).strip() + "\n")
                    CONSOLE.print(f"[*] Exported {len(names)} names to TXT: {save_path}", style="#00bb00")
                    messagebox.showinfo("导出成功", f"已导出到：{save_path}")
                except Exception as e:
                    CONSOLE.print(f"[!] Failed to export TXT: {e}", style="#bb0000")
                    messagebox.showerror("错误", f"导出TXT失败：{e}")

            def DeleteSelectedFile(self):
                filename = self.GetSelectedFile()
                if not filename:
                    return
                file_path = os.path.join(self.parent.choice_folder, filename)
                if messagebox.askyesno("确认删除", f"是否删除名单文件：{filename}?"):
                    try:
                        os.remove(file_path)
                        self.RefreshList()
                        messagebox.showinfo("删除成功", f"已删除：{filename}")
                    except Exception as e:
                        CONSOLE.print(f"[!] Failed to delete file: {e}", style="#bb0000")
                        messagebox.showerror("错误", f"删除文件失败：{e}")

            def CreateEmptyList(self):
                new_name = simpledialog.askstring("新建空名单", "请输入名单文件名（不含扩展名）：", parent=self.setroot)
                if not new_name:
                    return
                new_name = new_name.strip()
                if not new_name:
                    messagebox.showwarning("提示", "文件名不能为空。")
                    return
                if not new_name.lower().endswith(".json"):
                    new_name = f"{new_name}.json"
                target_path = os.path.join(self.parent.choice_folder, new_name)
                count = 1
                while os.path.exists(target_path):
                    base, ext = os.path.splitext(new_name)
                    target_path = os.path.join(self.parent.choice_folder, f"{base}_{count}{ext}")
                    count += 1
                try:
                    with open(target_path, "w", encoding="utf-8") as f:
                        json.dump([], f, ensure_ascii=False, indent=2)
                    self.RefreshList()
                    CONSOLE.print(f"[*] Created empty list file: {target_path}", style="#00bb00")
                    messagebox.showinfo("新建成功", f"已创建空名单：{os.path.basename(target_path)}")
                except Exception as e:
                    CONSOLE.print(f"[!] Failed to create file: {e}", style="#bb0000")
                    messagebox.showerror("错误", f"创建名单文件失败：{e}")

            def RenameSelectedFile(self):
                filename = self.GetSelectedFile()
                if not filename:
                    return
                old_path = os.path.join(self.parent.choice_folder, filename)
                new_name = simpledialog.askstring("重命名名单", "请输入新的文件名（不含扩展名）：", parent=self.setroot)
                if not new_name:
                    return
                new_name = new_name.strip()
                if not new_name:
                    messagebox.showwarning("提示", "文件名不能为空。")
                    return
                if not new_name.lower().endswith('.json'):
                    new_name = f"{new_name}.json"
                target_path = os.path.join(self.parent.choice_folder, new_name)
                if os.path.abspath(target_path) == os.path.abspath(old_path):
                    messagebox.showinfo("提示", "新旧文件名相同，无需重命名。")
                    return
                count = 1
                base, ext = os.path.splitext(new_name)
                while os.path.exists(target_path):
                    target_path = os.path.join(self.parent.choice_folder, f"{base}_{count}{ext}")
                    count += 1
                try:
                    os.rename(old_path, target_path)
                    if hasattr(self.parent, 'current_file_path') and os.path.abspath(old_path) == os.path.abspath(self.parent.current_file_path):
                        self.parent.current_file_path = os.path.abspath(target_path)
                        self.parent.UpdateLoadedFileLabel()
                        self.parent.SaveSetting('LoadFile', self.parent.current_file_path)
                    self.RefreshList()
                    CONSOLE.print(f"[*] File renamed: {filename} -> {os.path.basename(target_path)}", style="#00bb00")
                    messagebox.showinfo("重命名成功", f"已将 {filename} 重命名为 {os.path.basename(target_path)}")
                except Exception as e:
                    CONSOLE.print(f"[!] Failed to rename file: {e}", style="#bb0000")
                    messagebox.showerror("错误", f"重命名失败：{e}")

        setUI = SettingUI(self)

    def askFile(self):
        path = filedialog.askopenfilename(
            title="选择名单文件",
            filetypes=[("Text files", "*.txt"), ("JSON files", "*.json"), ("All files", "*")]
        )
        if not path:
            return
        self.LoadFile(path)

    def Choice(self):
        if not self.choiseList:
            self.showLabel.config(text="没有可抽取的选项", fg="red")
            return
        if self.rolling:
            return
        count = self.countVar.get()
        if count < 1:
            messagebox.showwarning("提示", "抽取个数至少为1")
            return
        if count > len(self.choiseList):
            messagebox.showwarning("提示", "抽取个数不能超过名单总数")
            return
        self.roll_target_count = count
        self.rolling = True
        self.startButton.state(["disabled"])
        self.showLabel.config(text="正在抽奖...", fg="blue")
        self.roll_start_time = time.time()
        self._roll_animation()

    def _roll_animation(self):
        if not self.rolling:
            return
        elapsed = time.time() - self.roll_start_time
        self.showLabel.config(text=random.choice(self.choiseList), fg="black")
        if elapsed < 3:
            self.roll_job = self.root.after(80, self._roll_animation)
            return
        self.rolling = False
        self.startButton.state(["!disabled"])
        winners = random.sample(self.choiseList, self.roll_target_count)
        result_text = "、".join(winners)
        self.showLabel.config(text=f"中奖：{result_text}", fg="green")

    def threadPrograss(self, *funcs):
        def run():
            for func in funcs:
                func()
        progress = threading.Thread(target=run)
        progress.daemon = True
        progress.start()

    @timer()
    def LoadSetting(self):
        try:
            if not os.path.exists("assets/setting.json"):
                default_setting = {"ScreenPos": {"x": 225, "y": 231}, "LoadFile": None}
                with open("assets/setting.json", "w", encoding="utf-8") as f:
                    json.dump(default_setting, f, ensure_ascii=False, indent=2)
            with open("assets/setting.json", "r", encoding="utf-8") as f:
                setting = json.load(f)
                self.latestPos = setting.get("ScreenPos", {"x": 225, "y": 231})
                self.root.geometry(f"800x300+{self.latestPos['x']}+{self.latestPos['y']}")
                load_file = setting.get("LoadFile")
                if load_file:
                    self.LoadFile(load_file)
        except Exception as e:
            CONSOLE.print(f"[!] Failed to load settings: {e}", style="#bb0000")
            messagebox.showerror("Error", f"Failed to load settings: {e}")

    @timer()
    def SaveSetting(self, key, value):
        try:
            if not os.path.exists("assets/setting.json"):
                setting = {"ScreenPos": self.latestPos, "LoadFile": None}
            else:
                with open("assets/setting.json", "r", encoding="utf-8") as f:
                    setting = json.load(f)
            setting[key] = value
            with open("assets/setting.json", "w", encoding="utf-8") as f:
                json.dump(setting, f, ensure_ascii=False, indent=2)

            CONSOLE.print(f"[*] Setting saved: {key} = {value}", style="#00bb00")
        except Exception as e:
            CONSOLE.print(f"[!] Failed to save setting: {e}", style="#bb0000")
            messagebox.showerror("Error", f"Failed to save setting: {e}")

    def LoadChoiceFiles(self):
        files = []
        try:
            for name in os.listdir(self.choice_folder):
                if name.lower().endswith(".json"):
                    files.append(name)
        except Exception:
            pass
        return files

    def UpdateCountLimit(self):
        total = len(self.choiseList)
        maxval = max(1, min(100, total))
        self.countSpinBox.config(to=maxval)
        if self.countVar.get() > maxval:
            self.countVar.set(maxval)
        self.countInfoLabel.config(text=f"名单数量：{total}")
        self.UpdateLoadedFileLabel()
        if total == 0:
            self.startButton.state(["disabled"])
        else:
            self.startButton.state(["!disabled"])

    def UpdateLoadedFileLabel(self):
        if self.current_file_path:
            self.loadedFileLabel.config(text=f"当前名单：{os.path.basename(self.current_file_path)}")
        else:
            self.loadedFileLabel.config(text="当前名单：无")

    @timer()
    def LoadFile(self, file_path=None):
        if not file_path:
            print('None File.')
            return
        try:
            if file_path.lower().endswith(".txt"):
                self.ImportTxtFile(file_path)
                return
            if file_path.lower().endswith(".json"):
                with open(file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if not isinstance(data, list):
                    raise ValueError("名单文件格式不正确，应为名称列表。")
                self.choiseList = self.NormalizeNames(data)
                if not self.choiseList:
                    messagebox.showwarning("提示", "名单文件为空。")
                self.current_file_path = file_path
                self.UpdateCountLimit()
                self.SaveSetting("LoadFile", file_path)
                return
            messagebox.showwarning("提示", "仅支持 TXT 或 JSON 文件导入。")
        except Exception as e:
            CONSOLE.print(f"[!] Failed to load file: {e}", style="#bb0000")
            messagebox.showerror("Error", f"Failed to load file: {e}")
            return

    @timer()
    def ImportTxtFile(self, file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                lines = [line.strip() for line in f if line.strip()]
            if not lines:
                messagebox.showwarning("提示", "文本文件没有可用名单内容。")
                return
            names = self.NormalizeNames(lines)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            target_name = base_name + ".json"
            target_path = os.path.join(self.choice_folder, target_name)
            count = 1
            while os.path.exists(target_path):
                target_name = f"{base_name}_{count}.json"
                target_path = os.path.join(self.choice_folder, target_name)
                count += 1
            with open(target_path, "w", encoding="utf-8") as f:
                json.dump(names, f, ensure_ascii=False, indent=2)
            self.choiseList = names
            self.current_file_path = target_path
            self.UpdateCountLimit()
            self.SaveSetting("LoadFile", target_path)
            message = f"已导入 {len(names)} 条名单，并保存为：{target_path}"
            if len(names) < len(lines):
                message += f"\n已自动去重，去除重复项 {len(lines) - len(names)} 条。"
            messagebox.showinfo("导入成功", message)
        except Exception as e:
            CONSOLE.print(f"[!] Failed to import TXT: {e}", style="#bb0000")
            messagebox.showerror("Error", f"Failed to import TXT: {e}")
            return

    @timer()
    def LoadNamesFromJson(self, file_path):
        if not file_path:
            raise ValueError("无效的名单文件路径。")
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            raise ValueError("名单文件格式不正确，应为名称列表。")
        return self.NormalizeNames(data)

    def NormalizeNames(self, names):
        seen = set()
        normalized = []
        for item in names:
            name = str(item).strip()
            if not name or name in seen:
                continue
            seen.add(name)
            normalized.append(name)
        return normalized

    def SaveNamesToJson(self, file_path, names):
        if not file_path:
            raise ValueError("目标文件路径无效，无法保存名单。")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(names, f, ensure_ascii=False, indent=2)

    def OpenFileEditor(self, file_path):
        FileEditor(self, file_path)

    def UIposChanged(self):
        while True:
            if self.WindowMove():
                x = self.root.winfo_x()
                y = self.root.winfo_y()
                if self.latestPos!= {"x": x, "y": y}:
                    self.SaveSetting("ScreenPos", {"x": x, "y": y})
                    self.latestPos = {"x": x, "y": y}
                CONSOLE.print("[*] Window position changed: ", self.latestPos, style="#00bb00")
            time.sleep(0.1)

    def tr(self, text):
        return text
    
    def WindowMove(self):
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        if {"x": x, "y": y} != self.latestPos:
            return True
        return False

class FileEditor:
    def __init__(self, parent, file_path):
        self.parent = parent
        if not file_path:
            messagebox.showerror("错误", "未提供名单文件路径，无法打开编辑器。")
            return
        self.file_path = os.path.abspath(file_path)
        self.names = self.parent.LoadNamesFromJson(self.file_path)

        self.editor = tk.Toplevel(parent.root)
        self.editor.title(f"编辑名单 - {os.path.basename(file_path)}")
        self.editor.resizable(False, False)
        self.editor.iconbitmap("assets/icon.ico")
        self.editor.geometry("420x380")
        self.SetControls()
        self.RefreshNames()
        self.editor.grab_set()

    def SetControls(self):
        title_label = tk.Label(self.editor, text="名单内容", font=(self.parent.font, 12))
        title_label.place(x=20, y=10)

        self.nameListbox = tk.Listbox(self.editor, height=14, width=40)
        self.nameListbox.place(x=20, y=40)

        self.addEntry = ttk.Entry(self.editor, width=30)
        self.addEntry.place(x=20, y=320)

        self.addButton = ttk.Button(self.editor, text="添加", command=self.AddName)
        self.addButton.place(x=300, y=318, width=80)

        self.deleteButton = ttk.Button(self.editor, text="删除选中", command=self.DeleteSelected)
        self.deleteButton.place(x=20, y=350, width=120)

        self.saveButton = ttk.Button(self.editor, text="保存修改", command=self.SaveNames)
        self.saveButton.place(x=150, y=350, width=120)

        self.closeButton = ttk.Button(self.editor, text="关闭", command=self.editor.destroy)
        self.closeButton.place(x=280, y=350, width=120)

        self.addrange = ttk.Button(self.editor, text="添加序列", command=self.ImportRange)
        self.addrange.place(x=310, y=280, width=80)

    def ImportRange(self):
        startnum = simpledialog.askinteger("添加序列", "请输入起始编号：", parent=self.editor)
        endnum = simpledialog.askinteger("添加序列", "请输入结束编号：", parent=self.editor)
        if startnum is None or endnum is None:
            return
        if startnum > endnum:
            messagebox.showwarning("提示", "起始编号必须小于等于结束编号。")
            return
        for i in range(startnum, endnum+1):
            name = str(i).strip()
            if not name:
                continue
            if name in self.names:
                continue
            self.names.append(name)

        self.RefreshNames()

    def RefreshNames(self):
        self.nameListbox.delete(0, tk.END)
        for name in self.names:
            self.nameListbox.insert(tk.END, name)
        CONSOLE.print(f"[*] Names refreshed: {len(self.names)}", style="#00bb00")

    def AddName(self):
        new_name = self.addEntry.get().strip()
        if not new_name:
            CONSOLE.print("[-] No name provided for addition.", style="#bbbb00")
            messagebox.showwarning("提示", "请输入要添加的姓名。")
            return
        self.names.append(new_name)
        self.addEntry.delete(0, tk.END)
        self.RefreshNames()

    def DeleteSelected(self):
        selection = self.nameListbox.curselection()
        if not selection:
            CONSOLE.print("[-] No name selected for deletion.", style="#bbbb00")
            messagebox.showwarning("提示", "请先选择一个要删除的姓名。")
            return
        for index in reversed(selection):
            self.names.pop(index)
        self.RefreshNames()

    def SaveNames(self):
        if not self.file_path:
            CONSOLE.print("[-] No file path provided, cannot save names.", style="#bbbb00")
            messagebox.showerror("错误", "当前名单文件路径无效，请重新打开名单文件后再保存。")
            return
        try:
            self.parent.SaveNamesToJson(self.file_path, self.names)
            if hasattr(self.parent, 'current_file_path') and os.path.abspath(self.file_path) == os.path.abspath(self.parent.current_file_path):
                self.parent.choiseList = self.names
                self.parent.UpdateCountLimit()
            CONSOLE.print(f"[*] Names saved: {len(self.names)}", style="#00bb00")
            messagebox.showinfo("保存成功", "名单已保存。")
        except Exception as e:
            CONSOLE.print(f"[!] Failed to save names: {e}", style="#bb0000")
            messagebox.showerror("错误", f"保存名单失败：{e}")

if __name__ == "__main__":
    app = MainUI()