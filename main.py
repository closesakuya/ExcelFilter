import json
import os
import re
import sys
import time
import threading
from PySide2.QtCore import Slot
from PySide2.QtGui import QTextCursor, QIcon, QPixmap
from PySide2.QtWidgets import QApplication, QMainWindow, QLineEdit, QTextEdit, \
    QPushButton, \
    QFileDialog
from PySide2.QtCore import Signal, Slot, QDateTime, QTimer, QEventLoop, QCoreApplication
from main_ui import Ui_water_mainwd
from Filter import FilterType, get_input_tab_by_filename, get_output_tab_by_filename, Task
import imgs


class UI(QMainWindow, Ui_water_mainwd):
    signal_log = Signal(str, bool, object)

    def __init__(self, *wd, **kw):
        Ui_water_mainwd.__init__(self)
        QMainWindow.__init__(self, parent=None)
        self.setupUi(self)
        icon = QIcon()
        icon.addPixmap(QPixmap(":res/main.ico"), QIcon.Normal, QIcon.Off)
        self.setWindowIcon(icon)

        # 所有tool button统一注册
        for k, item in self.__dict__.items():
            if isinstance(item, QPushButton):
                self.__getattribute__(k).clicked.connect(self.on_btn_clicked)
        self.dir_selector_map = {
            self.sel_output_btn: self.sel_output_lbl
        }
        self.file_selector_map = {
            self.sel_input_btn: self.sel_input_lbl
        }
        self.file_exec_map = {
            self.load_setting_btn: self.load_setting,
            self.dump_setting_btn: self.dump_setting
        }

        self.signal_log.connect(self._log_msg)
        # self.clear_input_btn.hide()

        self.filter_num = 0
        for i in range(100):
            if hasattr(self, "filter_fr_{0}".format(i + 1)):
                self.filter_num = i + 1



        self.load_input_set()

        self.__task_map = {}
        self.__task_lock = threading.RLock()

        self.start()

    def dump_input_set(self):
        dct = {}
        for k, item in self.__dict__.items():
            if isinstance(item, QLineEdit):
                dct[k] = item.text()
            elif isinstance(item, QTextEdit):
                dct[k] = item.toPlainText()
        with open(".ui.dump", "w+", encoding="utf-8") as f:
            f.write(json.dumps(dct, indent=1, ensure_ascii=False))

    def load_input_set(self):
        try:
            with open(".ui.dump", "r", encoding="utf-8") as f:
                dct = json.load(f)
                if dct:
                    for k, v in dct.items():
                        if hasattr(self, k):
                            item = self.__getattribute__(k)
                            if isinstance(item, QLineEdit) or isinstance(item, QTextEdit):
                                item.setText(v)
        except FileNotFoundError:
            pass
        except Exception as e:
            print(e)

    def closeEvent(self, event):
        print("main window close event")
        self.dump_input_set()
        # os._exit(0)
        sys.exit(0)

    @Slot()
    def on_btn_clicked(self):
        obj = self.sender()
        # 文件选择器
        if self.file_selector_map.get(obj, None) is not None:
            return self.on_common_file_choice_btn_clicked(obj, sel="file",
                                                          txt_show=self.file_selector_map.get(obj, None))

        # 文件夹选择器
        elif self.dir_selector_map.get(obj, None) is not None:
            return self.on_common_file_choice_btn_clicked(obj, sel="dir",
                                                          txt_show=self.dir_selector_map.get(obj, None))
        # 文件选择后执行回调
        elif self.file_exec_map.get(obj, None) is not None:
            return self.on_common_file_choice_btn_clicked(obj, sel="file", callback=self.file_exec_map[obj])

    def on_common_file_choice_btn_clicked(self, obj, file_filter="*.*", sel="file", txt_show=None, callback=None):
        if txt_show is not None:
            curpath = txt_show.text()
            curpath = os.path.split(curpath)[0]
            curpath = curpath if curpath else "."
        else:
            curpath = ""
        if "file" == sel:
            # file_name = QFileDialog.getOpenFileName(None, "文件选择", curpath, file_filter)
            file_name = QFileDialog.getOpenFileNames(None, "文件选择", curpath, file_filter)
        else:
            file_name = [QFileDialog.getExistingDirectory(None, "文件选择", curpath)]
        show_text = ""
        if isinstance(file_name, tuple):
            if isinstance(file_name[0], list):
                show_text = "||".join(file_name[0])
            else:
                show_text = file_name[0]
        elif isinstance(file_name, list):
            show_text = file_name[0]
        if txt_show:
            txt_show.setText(show_text)

        if callable(callback):
            for item in show_text.split("||"):
                callback(item)

    def load_setting(self, path):
        pass
        # try:
        #     with open(path, "r", encoding="utf-8") as f:
        #         dct = json.load(f)
        #         if dct:
        #             # 先清空
        #             for i in range(99):
        #                 if hasattr(self, "{0}_{1}".format("item_name", i + 1)):
        #                     self.__getattribute__("{0}_{1}".format("item_name", i + 1)).setText("")
        #                     self.__getattribute__("{0}_{1}".format("item_reg", i + 1)).setText("")
        #                     self.__getattribute__("{0}_{1}".format("item_index", i + 1)).setText("0")
        #                     self.__getattribute__("{0}_{1}".format("item_skip", i + 1)).setText("0")
        #                 else:
        #                     break
        #             for k, v in dct.items():
        #                 if hasattr(self, k):
        #                     item = self.__getattribute__(k)
        #                     if isinstance(item, QLineEdit) or isinstance(item, QTextEdit):
        #                         item.setText(v)
        #             # 更新读取字典
        #             if dct.get("items_lists", None) is not None and isinstance(dct["items_lists"], list):
        #                 for idx, item in enumerate(dct["items_lists"]):
        #                     if not isinstance(item, dict):
        #                         self.log_msg(u"{0} 读取异常".format(item))
        #                         continue
        #                     for k_name in ["item_name", "item_reg", "item_index", "item_skip"]:
        #                         if hasattr(self, "{0}_{1}".format(k_name, idx + 1)):
        #                             self.__getattribute__("{0}_{1}".format(k_name, idx + 1)).setText(
        #                                 str(item.get(k_name, "")))
        #             else:
        #                 self.log_msg(u"未发现items_lists 或其不为列表")
        #
        #         self.log_msg(u"从:{0} 读取配置成功".format(path))
        # except Exception as e:
        #     self.log_msg(str(e))


    def dump_setting(self, path):
        pass
        # try:
        #     with open(path, "w+", encoding="utf-8") as f:
        #         dct = {}
        #         for k, item in self.__dict__.items():
        #             if isinstance(item, QLineEdit):
        #                 if k.endswith("_filter"):
        #                     dct[k] = item.text()
        #         dct["items_lists"] = []
        #         idx_cnt = 1
        #         while 1:
        #             if not hasattr(self, "{0}_{1}".format("item_name", idx_cnt)):
        #                 break
        #             per_line = {}
        #             for k_name in ["item_name", "item_reg", "item_index", "item_skip"]:
        #                 if hasattr(self, "{0}_{1}".format(k_name, idx_cnt)):
        #                     per_line[k_name] = self.__getattribute__("{0}_{1}".format(k_name, idx_cnt)).text()
        #             dct["items_lists"].append(per_line)
        #             idx_cnt += 1
        #         f.write(json.dumps(dct, indent=1, ensure_ascii=False))
        #         self.log_msg(u"保存配置文件到:{0} 成功".format(path))
        # except Exception as e:
        #     self.log_msg(str(e))

    def log_msg(self, msg, mv_end=True, replace_pattern: str or re.RegexFlag = ""):
        self.signal_log.emit(msg, mv_end, replace_pattern)

    def _log_msg(self, msg, mv_end=True, replace_pattern: str or re.RegexFlag = ""):
        cursor = self.result_lbl.textCursor()
        if replace_pattern:
            ret = []
            found = False
            for item in self.result_lbl.toPlainText().split("\n"):
                if re.search(replace_pattern, item):
                    ret.append(msg)
                    found = True
                else:
                    ret.append(item)
            if not found:
                ret.append(msg)
            self.result_lbl.setText("\n".join(ret))
        else:
            cursor.insertText(msg + "\n")

        if mv_end:
            cursor.movePosition(QTextCursor.End)
            self.result_lbl.setTextCursor(cursor)

    @Slot()
    def on_clear_input_btn_clicked(self):
        for k, item in self.__dict__.items():
            if isinstance(item, QLineEdit):
                if k.startswith("item_"):
                    item.clear()

    @Slot()
    def on_clear_output_btn_clicked(self):
        self.result_lbl.clear()

    @Slot()
    def on_clear_input_btn_clicked(self):
        for i in range(self.filter_num):
            self.__getattribute__("filter_input_{0}".format(i+1)).clear()
            # self.__getattribute__("title_start_row_{0}".format(i + 1)).clear()

    @Slot()
    def on_exec_btn_clicked(self):
        self.clear_output_btn.click()
        self.log_msg("请等待...")
        output_file_name = self.sel_output_file.text()
        if not os.path.splitext(output_file_name)[-1]:
            output_file_name = output_file_name + '.xlsx'
        if not self.sel_output_lbl.text():
            dst = output_file_name
        else:
            dst = self.sel_output_lbl.text() + os.sep + output_file_name
        src = self.sel_input_lbl.text()
        try:
            ai = get_input_tab_by_filename(src)(
                src, data_pos=(int(self.data_start_row.text()) - 1, int(self.data_start_column.text()) - 1),
                title_pos=(int(self.title_start_row.text()) - 1, int(self.title_start_column.text()) - 1))
        except Exception as e:
            self.log_msg(e.__str__())
            return
        for i in range(self.filter_num):
            filter_txt = self.__getattribute__("filter_input_{0}".format(i + 1)).toPlainText().strip()
            if filter_txt:
                if self.__getattribute__("sel_btn_lst_{0}".format(i+1)).isChecked():
                    filter_type = FilterType.raw_list
                elif self.__getattribute__("sel_btn_py_{0}".format(i+1)).isChecked():
                    filter_type = FilterType.py_exp
                else:
                    filter_type = FilterType.reg_exp
                filter_col = self.__getattribute__("title_start_row_{0}".format(i+1)).text()
                print(filter_type, filter_txt, filter_col)
                ai.add_filter(filter_type, filter_txt, filter_col)

        title = ai.read_title()
        filter_title = self.filter_result_items.toPlainText().strip().split('\n')
        ao = get_output_tab_by_filename(dst)(dst, title=title, filter_title=filter_title)
        tsk = Task(ai, ao)

        self.__task_lock.acquire()
        self.__task_map["{0}->{1}".format(
            os.path.split(src)[-1], os.path.split(dst)[-1])] = tsk
        self.__task_lock.release()
        st = time.time()
        tsk.start()
        while time.time() - st <= 120 and not tsk.is_done():
            QCoreApplication.processEvents(QEventLoop.AllEvents, 100)
            time.sleep(1)
        if tsk.is_done() and self.open_when_fin_btn.isChecked():
            try:
                os.system("explorer \\e,\\root,{0}".format(self.sel_output_lbl.text().replace("/", os.sep)))
            except Exception as e:
                print(e)

    def _routine(self):
        while True:
            self.__task_lock.acquire()
            for item in list(self.__task_map.items()):
                if item is not None:
                    k, v = item
                else:
                    continue
                if not isinstance(v, Task):
                    continue

                raw_pct = v.get_progress()
                raw_pct = raw_pct if raw_pct < 1 else 1
                pct = int(35 * raw_pct)
                p = "▋" * pct + " " * (35 - pct)
                self.log_msg("{0} 分析进度:{1} \t{2:.2f}%\n".format(k, p, 100*pct/35.0),
                             mv_end=True, replace_pattern="{0} 分析进度".format(k))
                if v.is_done():
                    if not v.fault_msg:
                        self.log_msg("{0} 分析完成，写入行数:{1}".format(k, v.get_write_len()))
                    else:
                        self.log_msg("发生错误 " + v.fault_msg + "\n")
                    self.__task_map.pop(k)

            self.__task_lock.release()
            QCoreApplication.processEvents(QEventLoop.AllEvents, 100)
            time.sleep(1)

    def start(self):
        t = threading.Thread(target=self._routine)
        t.setDaemon(True)
        t.start()


if __name__ == "__main__":
    uiapp = QApplication([])
    a = UI()
    a.show()
    sys.exit(uiapp.exec_())
