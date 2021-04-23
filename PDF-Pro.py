from os.path import split
from os import startfile
from elegantUI import *
from PySide2.QtWidgets import (QVBoxLayout, QHBoxLayout, QLabel,
                               QApplication, QWidget, QMainWindow, QLineEdit,
                               QFileDialog, QListWidgetItem, QStackedWidget,
                               QSpinBox, QRadioButton)
from PySide2.QtCore import Qt
from PyPDF2.merger import PdfFileMerger, PdfFileReader, PdfFileWriter
import tempfile
import win10toast

toast = win10toast.ToastNotifier()


def filter_name(name):
    return split(name)[-1]


class MergePDFWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.pdf_merger = PdfFileMerger(strict=False)
        self.box = QVBoxLayout(self)
        self.box.setContentsMargins(0, 20, 20, 10)
        self.construct_buttons()
        self.setLayout(self.box)

    def construct_buttons(self):
        self.file_path = QLineEdit(self)
        self.file_path.setPlaceholderText('这里会显示文件路径')
        self.file_path.setReadOnly(True)
        self.file_path.setStyleSheet('padding: 10px;')
        # self.btn_select_file = buttons.DarkerButton(icon='打开文件.png', text='添加PDF',
        #                                          parent=self, on_press=self.btn_add_file_clicked)
        self.pdf_listview = lists.SmartList(self,
                                            on_item_click=self.pdf_item_clicked,
                                            on_item_double_click=self.pdf_item_double_clicked,
                                            add_filter=filter_name)
        self.btn_select_file = self.pdf_listview.add_btn_add('添加PDF', 'darker','打开文件.png',  self.btn_add_file_clicked)
        self.btn_delete_item = self.pdf_listview.add_delete_item_btn('删除选定项目',
                                                                     btn_type='darker',
                                                                     icon='删除一项.png',
                                                                     delete_callback=self.del_callback)
        self.btn_clear_all = self.pdf_listview.add_clear_btn('清空所有项目',
                                                             btn_type='darker',
                                                             icon='删除.png',
                                                             clear_callback=self.clear_callback)
        self.btn_download = buttons.DarkerButton(icon='下载文件.png', text='导出PDF文件',
                                                 parent=self, on_press=self.merge_and_write)
        self.btn_download.setDisabled(True)
        self.first_line_hbox = QHBoxLayout(self)
        self.first_line_hbox.addWidget(self.file_path)
        self.first_line_hbox.addWidget(self.btn_select_file)
        self.box.addLayout(self.first_line_hbox)
        self.box.addWidget(self.pdf_listview)
        self.third_line_hbox = QHBoxLayout(self)
        self.third_line_hbox.addWidget(self.btn_delete_item)
        self.third_line_hbox.addWidget(self.btn_clear_all)
        self.box.addLayout(self.third_line_hbox)
        self.box.addWidget(self.btn_download)

    def del_callback(self):
        if len(self.pdf_listview.items) == 0:
            self.btn_download.setDisabled(True)

    def clear_callback(self):
        if len(self.pdf_listview.items) == 0:
            self.btn_download.setDisabled(True)

    def pdf_item_clicked(self, item: QListWidgetItem):
        """When the user clicked an item in pdf_listview(QListWidget), changes the
        file_path(QLineEdit)'s text to the context of the clicked item."""
        self.file_path.setText(self.pdf_listview.items[self.pdf_listview.currentIndex().row()])
        self.btn_delete_item.setDisabled(False)

    def pdf_item_double_clicked(self, item: QListWidgetItem):
        if dialogs.Messages().send_question('打开', '是否打开这个PDF文件？') == 0:
            startfile(self.pdf_listview.items[self.pdf_listview.currentIndex().row()])

    def btn_add_file_clicked(self):
        filename, _ = QFileDialog.getOpenFileName(self, '添加文件', '', 'PDF文件(*.pdf)')
        if not filename:
            return
        self.pdf_listview.addItem(filename)
        self.file_path.setText(filename)
        self.btn_clear_all.setDisabled(False)
        self.btn_download.setDisabled(False)

    def merge_and_write(self):
        filename, _ = QFileDialog.getSaveFileName(self, '保存PDF', '', 'PDF文件(*.pdf)')
        if not filename:
            return
        merger = PdfFileMerger(strict=False)
        
        try:
            self.btn_download.setText('正在导出······')
            self.btn_download.setDisabled(True)
            for item in self.pdf_listview.items:
                while True:
                    try:
                        merger.append(open(item, 'rb'))
                    except (OSError, Exception):
                        information = 'PDF文件无法打开，也许是因为格式不正确，也可能是正在被其他程序使用。' \
                                      '请关掉可能使用它的程序后再试。'
                        toast.show_toast('PDF文件无法打开', information, threaded=True)
                    else:
                        break
            merger.write(open(filename, 'wb'))
            self.btn_download.setText('导出PDF文件')
            self.btn_download.setDisabled(False)
            toast.show_toast('成功', '成功合并PDF，已导出！', threaded=True)
        except (IOError, OSError):
            toast.show_toast('错误', '由于权限错误，无法导出PDF，请换一个路径再试。', threaded=True)
            return


class ExtractPDFWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.master = parent
        self.selected = None
        self.construct_UI()

    def filter_name(self, labels):
        result = f'{split(labels[0])[-1]}中的第 {labels[1]} 至 {labels[2]} 页'
        return result

    def construct_UI(self):
        self.vbox = QVBoxLayout(self)
        self.vbox.setContentsMargins(0, 20, 20, 10)
        # setup the first line of UI
        self.first_line_hbox = QHBoxLayout()
        self.first_line_hbox.setContentsMargins(0, 0, 0, 0)
        self.line_edit_file_path = QLineEdit(self)
        self.line_edit_file_path.setPlaceholderText('这里会显示文件路径')
        self.line_edit_file_path.setReadOnly(True)
        self.line_edit_file_path.setStyleSheet('padding: 10px;')
        self.first_line_hbox.addWidget(self.line_edit_file_path)
        self.list = lists.SmartList(self, on_item_click=self.on_list_item_selected,
                                    on_item_double_click=self.on_double_click, add_filter=self.filter_name)
        self.btn_add_pdf = self.list.add_btn_add('选择PDF', 'darker', '打开文件.png', self.add_pdf_dialog_triggered)
        self.first_line_hbox.addWidget(self.btn_add_pdf)
        self.vbox.addLayout(self.first_line_hbox)
        self.second_line_hbox = QHBoxLayout()
        self.hbox3 = QHBoxLayout()
        self.spin_start = spinbox.SpinBox(self)
        self.spin_start.setMinimum(1)
        self.spin_start.setDisabled(True)
        self.spin_start.valueChanged.connect(self.refresh)
        self.spin_end = spinbox.SpinBox(self)
        self.spin_end.setMinimum(1)
        self.spin_end.setDisabled(True)
        self.spin_end.valueChanged.connect(self.refresh)
        self.btn_submit = buttons.DarkerButton(text='添加已选择的PDF文件', parent=self,
                                               on_press=self.on_add, icon='添加.png')
        self.btn_submit.setDisabled(True)
        l = QLabel('抽取')
        l.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.hbox3.addWidget(l)
        self.hbox3.addWidget(self.spin_start)
        self.hbox3.addWidget(QLabel('至'))
        self.hbox3.addWidget(self.spin_end)
        self.hbox3.addWidget(QLabel('页'))
        self.second_line_hbox.addLayout(self.hbox3)
        self.second_line_hbox.addWidget(self.btn_submit)
        self.vbox.addLayout(self.second_line_hbox)

        # setup the third line of UI

        self.vbox.addWidget(self.list)

        # setup the fourth line of UI
        self.fourth_line_hbox = QHBoxLayout(self)
        self.fourth_line_hbox.setContentsMargins(0, 0, 0, 0)
        self.btn_del_item = self.list.add_delete_item_btn('删除选定项目', 'darker', '删除一项.png', self.del_callback)
        self.btn_del_item.setDisabled(True)
        self.btn_clear = self.list.add_clear_btn('删除所有项目', 'darker', '删除.png', self.clear_callback)
        self.btn_clear.setDisabled(True)
        self.fourth_line_hbox.addWidget(self.btn_del_item)
        self.fourth_line_hbox.addWidget(self.btn_clear)
        self.vbox.addLayout(self.fourth_line_hbox)

        self.btn_export = buttons.DarkerButton('导出PDF文件', self.on_export, self, icon='下载文件.png')
        self.btn_export.setDisabled(True)
        self.vbox.addWidget(self.btn_export)
        self.setLayout(self.vbox)

    def add_pdf_dialog_triggered(self):
        filename, _ = QFileDialog.getOpenFileName(self, '添加PDF', filter='PDF文件(*.pdf)')
        if filename:
            try:
                pdf = PdfFileReader(open(filename, 'rb'), strict=False)
                max_pages = pdf.numPages
                self.master.statusBar().showMessage(f'这个PDF文件共有{max_pages}页。', 10000)
                self.spin_end.setMaximum(max_pages)
                self.spin_start.setMaximum(max_pages)
            except OSError:
                information = 'PDF文件无法打开，也许是因为格式不正确，也可能是正在被其他程序使用。' \
                              '请关掉可能使用它的程序后再试。'
                toast.show_toast('PDF文件无法打开', information, threaded=True)
                return
            else:
                self.spin_start.setDisabled(False)
                self.spin_end.setDisabled(False)
                self.btn_submit.setDisabled(False)
                self.selected = filename
                self.line_edit_file_path.setText(filename)
                self.btn_add_pdf.setText('重新选择PDF')

    def del_callback(self):
        if len(self.list.items) == 0:
            self.btn_export.setDisabled(True)

    def clear_callback(self):
        if len(self.list.items) == 0:
            self.btn_export.setDisabled(True)

    def on_add(self):
        start = int(self.spin_start.value())
        end = int(self.spin_end.value())
        self.list.addItem((self.selected, start, end))
        self.btn_clear.setDisabled(False)
        self.btn_export.setDisabled(False)

    def refresh(self):
        if int(self.spin_start.value()) > int(self.spin_end.value()):
            self.btn_submit.setDisabled(True)
        else:
            self.btn_submit.setDisabled(False)

    def on_double_click(self):
        if dialogs.Messages().send_question('打开', '是否打开这一项的预览？') == 0:
            writer = PdfFileWriter()
            data = self.list.items[self.list.currentIndex().row()]
            filename = data[0]
            start, end = data[1], data[2]
            reader = PdfFileReader(open(filename, 'rb'))
            for page_num in range(int(start) - 1, int(end)):
                page = reader.getPage(page_num)
                writer.addPage(page)
            with tempfile.NamedTemporaryFile(delete=False, mode='wb', suffix='.pdf', prefix='预览_') as tmp:
                writer.write(tmp)
                startfile(tmp.name)

    def on_export(self):
        writer = PdfFileWriter()
        out, _ = QFileDialog.getSaveFileName(self, '导出', filter='PDF文件(*.pdf)')
        if not out:
            return
        self.btn_export.setDisabled(True)
        self.btn_export.setText('正在导出PDF······')
        for data in self.list.items:
            start, end = data[1], data[2]
            reader = PdfFileReader(open(data[0], 'rb'))
            for page_num in range(int(start) - 1, int(end)):
                page = reader.getPage(page_num)
                writer.addPage(page)
        with open(out, 'wb') as pdf:
            writer.write(pdf)
        toast.show_toast('成功', '成功抽取了PDF中的页码，已导出！', threaded=True)

    def on_list_item_selected(self):
        self.btn_del_item.setDisabled(False)


class DeletePDFWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected = None
        self.data = None
        self.master = parent
        self.construct_UI()

    def construct_UI(self):
        self.vbox = QVBoxLayout(self)
        self.btn_add = buttons.DarkerButton('选择PDF', icon='添加.png',
                                            parent=self, on_press=self.on_add_file)
        self.vbox.addWidget(self.btn_add)
        self.radio_btn_box = QHBoxLayout(self)
        self.radio_single_page = QRadioButton('删除单页', self)
        self.radio_single_page.setChecked(True)
        self.radio_single_page.toggled.connect(self.page_changed)
        self.radio_multi_page = QRadioButton('删除多页', self)
        self.radio_multi_page.toggled.connect(self.page_changed)
        self.radio_btn_box.addWidget(self.radio_single_page)
        self.radio_btn_box.addWidget(self.radio_multi_page)
        self.vbox.addLayout(self.radio_btn_box)
        self.stacked = QStackedWidget(self)
        self.hbox0 = QHBoxLayout(self)
        self.page = QSpinBox(self)
        self.page.valueChanged.connect(self.page_changed)
        self.page.setDisabled(True)
        self.hbox0.addWidget(QLabel('删除第'), alignment=Qt.AlignRight | Qt.AlignVCenter)
        self.hbox0.addWidget(self.page)
        self.hbox0.addWidget(QLabel('页'))
        w0 = QWidget()
        w0.setLayout(self.hbox0)
        self.stacked.addWidget(w0)
        self.hbox = QHBoxLayout(self)
        self.start = QSpinBox(self)
        self.end = QSpinBox(self)
        self.start.setDisabled(True)
        self.end.setDisabled(True)
        self.hbox.addWidget(QLabel('删除'), alignment=Qt.AlignRight | Qt.AlignVCenter)
        self.hbox.addWidget(self.start)
        self.hbox.addWidget(QLabel('至'))
        self.hbox.addWidget(self.end)
        self.hbox.addWidget(QLabel('的PDF页码'))
        w = QWidget()
        w.setLayout(self.hbox)
        self.stacked.addWidget(w)
        self.vbox.addWidget(self.stacked)
        self.btn_submit = buttons.DarkerButton('导出PDF', icon='下载文件.png',
                                               parent=self, on_press=self.on_export)
        self.btn_submit.setDisabled(True)
        self.vbox.addWidget(self.btn_submit)
        self.vbox.addStretch(0)
        self.setLayout(self.vbox)

    def on_add_file(self):
        filename, _ = QFileDialog.getOpenFileName(self, '添加文件', '', 'PDF文件(*.pdf)')
        if not filename:
            return
        try:
            pdf = PdfFileReader(open(filename, 'rb'))
            for i in [self.page, self.start, self.end]:
                i.setMinimum(1)
                i.setMaximum(pdf.numPages)
        except OSError:
            information = 'PDF文件无法打开，也许是因为格式不正确，也可能是正在被其他程序使用。' \
                          '请关掉可能使用它的程序后再试。'
            toast.show_toast('PDF文件无法打开', information, threaded=True)
            return
        else:
            self.selected = filename
            self.master.statusBar().showMessage('已选择：' + split(filename)[-1] + f'，共有{pdf.numPages}页。')
            self.page.setDisabled(False)
            self.start.setDisabled(False)
            self.end.setDisabled(False)

    def page_changed(self):
        if self.radio_single_page.isChecked():
            self.data = [int(self.page.value())]
            self.stacked.setCurrentIndex(0)
        elif self.radio_multi_page.isChecked():
            self.data = [int(self.start.value()), int(self.end.value())]
            self.stacked.setCurrentIndex(1)

    def on_export(self):
        save_filename = QFileDialog.getSaveFileName(
            self, '导出', '', 'PDF(*.pdf)')
        if not save_filename[0]:
            return
        if len(self.data) == 2:
            start = self.data[0]
            end = self.data[1]
        else:
            start = end = self.data[0]
        with open(self.selected, 'rb') as pdf:
            reader = PdfFileReader(pdf)
            with open(save_filename[0], 'wb') as save_pdf:
                writer = PdfFileWriter()
                for page in range(1, start):
                    writer.addPage(reader.getPage(page - 1))
                for page in range(end, reader.getNumPages()):
                    writer.addPage(reader.getPage(page))
                writer.write(save_pdf)
                toast.show_toast('成功', '成功删除PDF指定页码，已导出！', threaded=True)


class MainApplicationWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet('font-family: "Microsoft Yahei"')
        self.setWindowTitle('PDF Pro v3.1')
        self.cw = QWidget()
        self.cwl = QVBoxLayout(self.cw)
        self.main_tab = tabs.LightTab(self)
        self.main_tab.addTab(MergePDFWidget(), '合并PDF')
        self.main_tab.addTab(ExtractPDFWidget(self), '抽取PDF页码')
        self.main_tab.addTab(DeletePDFWidget(self), '删除PDF页码')
        self.cwl.addWidget(self.main_tab)
        self.cwl.setContentsMargins(10, 10, 0, 10)
        self.cw.setLayout(self.cwl)
        self.setCentralWidget(self.cw)


if __name__ == '__main__':
    app = QApplication()
    window = MainApplicationWindow()
    window.show()
    app.exec_()
