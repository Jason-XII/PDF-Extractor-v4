from os.path import split
from os import startfile
from JasonUI import *
from PySide2.QtWidgets import (QVBoxLayout, QHBoxLayout, QLabel,
                               QApplication, QWidget, QMainWindow, QLineEdit,
                               QFileDialog, QListWidgetItem, QStackedWidget,
                               QSpinBox, QRadioButton)
from PySide2.QtCore import Qt
from pdf_machine import *
import tempfile
import plyer.platforms.win.notification
from plyer import notification
import urllib.request


def filter_name(name):
    return split(name)[-1]


class MergePDFWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.construct_buttons()
        self.setLayout(self.box)

    def construct_buttons(self):
        self.file_path = QLineEdit(self)
        self.file_path.setPlaceholderText('这里会显示文件路径')
        self.file_path.setReadOnly(True)
        self.file_path.setStyleSheet('padding: 10px;')
        self.pdf_listview = lists.SmartList(self,
                                            on_item_click=self.pdf_item_clicked,
                                            on_item_double_click=self.pdf_item_double_clicked,
                                            add_filter=filter_name)
        self.btn_select_file = self.pdf_listview.add_btn_add(
            '添加PDF', 'darker', '打开文件.png',  self.btn_add_file_clicked)
        self.btn_move_up = self.pdf_listview.add_btn_up('向上移动项目', btn_type='darker', icon='up.png')
        self.btn_move_down = self.pdf_listview.add_btn_down('向下移动项目', 'darker', 'down.png')
        self.btn_delete_item = self.pdf_listview.add_delete_item_btn('删除选定项目',
                                                                     btn_type='darker',
                                                                     icon='删除一项.png',
                                                                     delete_callback=self.after_delete)
        self.btn_clear_all = self.pdf_listview.add_clear_btn('清空所有项目',
                                                             btn_type='darker',
                                                             icon='删除.png',
                                                             clear_callback=self.after_clear)
        self.btn_download = buttons.DarkerButton(icon='下载文件.png', text='导出PDF文件',
                                                 parent=self, on_press=self.merge_and_write)
        self.btn_download.setDisabled(True)
        self.first_line_hbox = layouts.HorizontalGroup(self.file_path, self.btn_select_file, parent=None)
        self.third_line_hbox = layouts.HorizontalGroup(self.btn_delete_item, self.btn_clear_all, self.btn_move_up, self.btn_move_down)
        self.box = layouts.VerticalGroup(self.first_line_hbox, self.pdf_listview, self.third_line_hbox, self.btn_download, parent=None)
        self.box.setContentsMargins(0, 20, 20, 10)

    def after_delete(self):
        if len(self.pdf_listview.items) == 0:
            self.btn_download.setDisabled(True)

    def after_clear(self):
        if len(self.pdf_listview.items) == 0:
            self.btn_download.setDisabled(True)

    def pdf_item_clicked(self, item: QListWidgetItem):
        """When the user clicked an item in pdf_listview(QListWidget), changes the
        file_path(QLineEdit)'s text to the context of the clicked item."""
        self.file_path.setText(
            self.pdf_listview.items[self.pdf_listview.currentIndex().row()])
        self.btn_delete_item.setDisabled(False)

    def pdf_item_double_clicked(self, item: QListWidgetItem):
        if dialogs.Messages().send_question('打开', '是否打开这个PDF文件？') == 0:
            startfile(
                self.pdf_listview.items[self.pdf_listview.currentIndex().row()])

    def btn_add_file_clicked(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, '添加文件', '', 'PDF文件(*.pdf)')
        if not filename:
            return
        self.pdf_listview.addItem(filename)
        self.file_path.setText(filename)
        self.btn_clear_all.setDisabled(False)
        self.btn_download.setDisabled(False)

    def merge_and_write(self):
        filename, _ = QFileDialog.getSaveFileName(
            self, '保存PDF', '', 'PDF文件(*.pdf)')
        if not filename:
            return
        merge_machine = PDFMergeMachine(self.pdf_listview.items)
        try:
            merge_machine.merge()
            notification.notify(
                title='成功', message='成功合并PDF，已导出！', app_icon='pdf-pro.ico')
        except (IOError, OSError):
            notification.notify(
                title='错误', message='由于未知错误，PDF导出失败！', app_icon='pdf-pro.ico')
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
        self.line_edit_file_path = QLineEdit(self)
        self.line_edit_file_path.setPlaceholderText('这里会显示文件路径')
        self.line_edit_file_path.setReadOnly(True)
        self.line_edit_file_path.setStyleSheet('padding: 10px;')
        self.list = lists.SmartList(self, on_item_click=self.on_list_item_selected,
                                    on_item_double_click=self.on_double_click, add_filter=self.filter_name)
        self.btn_add_pdf = self.list.add_btn_add(
            '选择PDF', 'darker', '打开文件.png', self.add_pdf_dialog_triggered)
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
        extract_label = QLabel('抽取')
        extract_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.btn_del_item = self.list.add_delete_item_btn(
            '删除选定项目', 'darker', '删除一项.png')
        self.btn_del_item.setDisabled(True)
        self.btn_clear = self.list.add_clear_btn(
            '删除所有项目', 'darker', '删除.png')
        self.btn_clear.setDisabled(True)
        self.btn_move_up = self.list.add_btn_up('向上移动项目', btn_type='darker', icon='up.png')
        self.btn_move_down = self.list.add_btn_down('向下移动项目', 'darker', 'down.png')
        self.btn_export = buttons.DarkerButton(
            '导出PDF文件', self.on_export, self, icon='下载文件.png')
        self.btn_export.setDisabled(True)
        self.first_line_hbox = layouts.HorizontalGroup(self.line_edit_file_path, self.btn_add_pdf)
        self.first_line_hbox.setContentsMargins(0, 0, 0, 0)
        self.inline_hbox = layouts.HorizontalGroup(extract_label, self.spin_start, QLabel('至'), self.spin_end, QLabel('页'))
        self.second_line_hbox = layouts.HorizontalGroup(self.inline_hbox, self.btn_submit)
        self.fourth_line_hbox = layouts.HorizontalGroup(self.btn_del_item, self.btn_clear, self.btn_move_up, self.btn_move_down)
        self.fourth_line_hbox.setContentsMargins(0, 0, 0, 0)
        self.vbox = layouts.VerticalGroup(self.first_line_hbox, self.second_line_hbox, self.list, self.fourth_line_hbox, self.btn_export)
        self.vbox.setContentsMargins(0, 20, 20, 10)
        self.setLayout(self.vbox)

    def add_pdf_dialog_triggered(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, '添加PDF', filter='PDF文件(*.pdf)')
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
                notification.notify('PDF文件无法打开', information,
                                    app_icon='pdf-pro.ico')
                return
            else:
                self.spin_start.setDisabled(False)
                self.spin_end.setDisabled(False)
                self.btn_submit.setDisabled(False)
                self.selected = filename
                self.line_edit_file_path.setText(filename)
                self.btn_add_pdf.setText('重新选择PDF')

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

        out, _ = QFileDialog.getSaveFileName(self, '导出', filter='PDF文件(*.pdf)')
        extract_machine = PDFExtractMachine(self.list.items)
        if not out:
            return
        extract_machine.extract_all(out)
        notification.notify(
            title='成功', message='成功抽取了PDF中的页码，已导出！', app_icon='pdf-pro.ico')

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
        self.hbox0.addWidget(
            QLabel('删除第'), alignment=Qt.AlignRight | Qt.AlignVCenter)
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
        self.hbox.addWidget(
            QLabel('删除'), alignment=Qt.AlignRight | Qt.AlignVCenter)
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
        filename, _ = QFileDialog.getOpenFileName(
            self, '添加文件', '', 'PDF文件(*.pdf)')
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
            notification.notify(title='PDF文件无法打开',
                                message=information, app_icon='pdf-pro.ico')
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
                notification.notify(
                    title='成功', message='成功删除PDF指定页码，已导出！', app_icon='pdf-pro')


class ExtractImageWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.btn_select_pdf = buttons.DarkerButton('添加PDF', self.on_add_pdf, self, '打开文件.png')
        self.btn_select_directory = buttons.DarkerButton('选择导出图片位置', self.on_add_directory, self, icon='打开文件.png')
        self.list = lists.SmartList(self, add_filter=lambda f: split(f)[-1])
        self.btn_del_item = self.list.add_delete_item_btn('删除选定项目', 'darker', '删除一项.png')
        self.btn_clear = self.list.add_clear_btn('清空所有项目', 'darker', '删除.png')
        self.btn_up = self.list.add_btn_up('向上移动项目', 'darker', 'up.png')
        self.btn_down = self.list.add_btn_down('向下移动项目', 'darker', 'down.png')
        self.btn_export = buttons.DarkerButton('导出图片至文件夹', self.on_extract, self, icon='下载文件.png')
        self.first_line_hbox = layouts.HorizontalGroup(self.btn_select_pdf, self.btn_select_directory)
        self.third_line_hbox = layouts.HorizontalGroup(self.btn_del_item, self.btn_clear, self.btn_up, self.btn_down)
        self.vbox = layouts.VerticalGroup(self.first_line_hbox, self.list, self.third_line_hbox, self.btn_export)
        self.setLayout(self.vbox)
        self.setContentsMargins(0, 20, 20, 10)

        self.dir = None

    def on_add_pdf(self):
        filename, _ = QFileDialog.getOpenFileName(self, '选择文件', filter='PDF文件(*.pdf)')
        if not filename:
            return
        try:
           open(filename, 'rb')
        except IOError:
            notification.notify('未知错误', '在打开文件时出现未知错误，无法抽取其中的图片。',
                                app_icon='pdf-pro.ico')
        else:
            self.list.addItem(filename)

    def on_add_directory(self):
        directory = QFileDialog.getExistingDirectory(self, '选择文件夹')
        self.dir = directory

    def on_extract(self):
        if self.dir is None:
            return
        extract_image_machine = PDFExtractImageMachine(self.list.items, self.dir)
        extract_image_machine.extract()
        notification.notify('成功', '抽取PDF图片成功，已导出！', app_icon='pdf-pro.ico')



class MainApplicationWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ver = "4.2.2"
        self.setStyleSheet('font-family: "Microsoft Yahei"')
        self.setWindowTitle(f'PDF Extractor {self.ver}')
        self.cw = QWidget()
        self.cwl = QVBoxLayout(self.cw)
        self.main_tab = tabs.LightTab(self)
        self.main_tab.addTab(MergePDFWidget(), '合并PDF')
        self.main_tab.addTab(ExtractPDFWidget(self), '抽取PDF页码')
        self.main_tab.addTab(ExtractImageWidget(self), '抽取PDF中的图片')
        self.cwl.addWidget(self.main_tab)
        self.cwl.setContentsMargins(10, 10, 0, 10)
        self.cw.setLayout(self.cwl)
        self.setCentralWidget(self.cw)
        self.create_menubar()

    def create_menubar(self):
        self.menu = self.menuBar()
        self.check_update_menu = self.menu.addMenu("检查更新")
        self.update_action = self.check_update_menu.addAction('检查更新')
        self.update_action.triggered.connect(self.is_update_available)

    def is_update_available(self):
        try:
            version_info = eval(urllib.request.urlopen(
                'http://jasoncoder16.pythonanywhere.com/version').read().decode())
            if self.ver < version_info['version']:
                notification.notify(
                    title='检测到更新', message=f'您现在的版本是{self.ver}, 但是PDF Extractor{version_info["version"]}已经正式发布。\
                    更新内容：{version_info["note"]}', app_icon='pdf-pro.ico')
            else:
                notification.notify(
                    title='不必更新', message='本产品已是最新版本。', app_icon='pdf-pro.ico')
        except Exception as err:
                
            raise
            notification.notify(title='错误', message='无法连接到服务器', app_icon='pdf-pro.ico')



if __name__ == '__main__':
    app = QApplication()
    window = MainApplicationWindow()
    window.show()
    app.exec_()
