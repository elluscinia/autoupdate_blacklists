'''

Модуль обработки пользовательского интерфейса

'''
import multiprocessing as mpro
import threading

import pythoncom
import win32com.client
from PyQt5 import QtCore, QtWidgets

import processSources
from interface import Ui_AutoUploadTMG


class Application(QtWidgets.QMainWindow, Ui_AutoUploadTMG, QtWidgets.QWidget):
    delete_DSN = QtCore.pyqtSignal(str)
    unpack_DSN = QtCore.pyqtSignal(str)
    import_DSN = QtCore.pyqtSignal(str)
    operation_DSN = QtCore.pyqtSignal(str)
    progressBar_fill = QtCore.pyqtSignal(int)
    progressBar_max = QtCore.pyqtSignal(int)
    progressBar_fill_download = QtCore.pyqtSignal(int)
    progressBar_max_download = QtCore.pyqtSignal(int)
    unblock_ParamArea = QtCore.pyqtSignal(bool)
    set_StatusBar = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        QtWidgets.QMainWindow.__init__(self, parent)
        self.setupUi(self)

        self.delete_DSN.connect(self.delete_clean_tmg)
        self.unpack_DSN.connect(self.add_unpack_tmg)
        self.import_DSN.connect(self.add_import_tmg)
        self.operation_DSN.connect(self.add_operation_tmg)
        self.progressBar_fill.connect(self.fill_progress_bar)
        self.progressBar_fill_download.connect(self.fill_progress_bar_download)
        self.progressBar_max.connect(self.set_progress_bar_max)
        self.progressBar_max_download.connect(self.set_progress_bar_max_download)
        self.unblock_ParamArea.connect(self.set_access_settings_and_commands)
        self.set_StatusBar.connect(self.set_status_bar)

        self.emit_dict = {
            # удаление из QListWidget domainNameSetsTMG наборов доменных имён
            'clean_GUI': self.delete_DSN,

            # добавление в QListWidget domainNameSetsMemory наборов доменных имён
            'unpack_GUI': self.unpack_DSN,

            # удаление файлов из domainNameSetsMemory и добавление в domainNameSetsTMG
            'import_GUI': self.import_DSN,

            # запись логов в QListWidget operations
            'operation_GUI': self.operation_DSN,

            # заполнение прогресс бара progressBar
            'fill_progressBar': self.progressBar_fill,

            # заполнение прогресс бара progressBar_download
            'fill_progressBar_download': self.progressBar_fill_download,

            # установка максимального значения в прогресс бар progressBar
            'progressBar_max': self.progressBar_max,

            # установка максимального значения в прогресс бар progressBar_download
            'progressBar_max_download': self.progressBar_max_download,

            # разблокировка области параметров и меню бара
            'unblock_area': self.unblock_ParamArea,

            # сделать запись в статус бар
            'set_StatusBar': self.set_StatusBar

        }

        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowSystemMenuHint)

        self.progressBar.setValue(0)
        self.progressBar_download.setValue(0)

        self.action_import.setEnabled(False)
        self.action_clean.setEnabled(False)

        try:
            pythoncom.CoInitialize()
            object_tmg = win32com.client.Dispatch('FPC.Root')
            isa_array = object_tmg.GetContainingArray()
        except pythoncom.pywintypes.com_error as ce:
            self.statusBar.showMessage('Отсутствует подключение к Forefront TMG', -1)
            self.set_access_settings_and_commands(False)
            return
        else:
            self.choice_ruleName.clear()
            self.choice_ruleName.addItems([isa_array.ArrayPolicy.PolicyRules.Item(i).Name
                                           for i in range(1, isa_array.ArrayPolicy.PolicyRules.Count + 1)])
            self.choice_ruleName.lineEdit().setAlignment(QtCore.Qt.AlignCenter)

            self.refresh_information()

    @QtCore.pyqtSlot(str)
    def delete_clean_tmg(self, data):
        '''
        Метод для удаления наборов доменных имён из виджета с текущим состоянием правила
        :param data: имя набора
        :return:
        '''
        list_items = self.domainNameSetsTMG.findItems(data, QtCore.Qt.MatchExactly)
        if not list_items:
            return
        for item in list_items:
            self.domainNameSetsTMG.takeItem(self.domainNameSetsTMG.row(item))

    @QtCore.pyqtSlot(str)
    def add_unpack_tmg(self, data):
        '''
        Метод для добавления наборов доменных имён в виджет с текущими обрабатываемыми XML
        :param data: имя набора
        :return:
        '''
        item = QtWidgets.QListWidgetItem(data)
        self.domainNameSetsMemory.addItem(item)

    @QtCore.pyqtSlot(str)
    def add_import_tmg(self, data):
        '''
        Метод для удаления наборов из виджета с текущим состоянием памяти и добавления в виджет с наборами в правиле
        :param data: имя набора
        :return:
        '''
        item = QtWidgets.QListWidgetItem(data)
        list_items = self.domainNameSetsMemory.findItems(data, QtCore.Qt.MatchExactly)
        if not list_items:
            return
        for item in list_items:
            self.domainNameSetsMemory.takeItem(self.domainNameSetsMemory.row(item))
        self.domainNameSetsTMG.addItem(item)

    @QtCore.pyqtSlot(str)
    def add_operation_tmg(self, data):
        '''
        Метод добавления операции в лог
        :param data: описание операции
        :return:
        '''
        item = QtWidgets.QListWidgetItem(data)
        self.operations.addItem(item)
        self.operations.scrollToBottom()

    @QtCore.pyqtSlot(int)
    def fill_progress_bar(self, value):
        '''
        Загрузка прогресс бара импорта
        :param value: инкремент
        :return:
        '''
        prev = self.progressBar.value()
        self.progressBar.setValue(prev + value)

    @QtCore.pyqtSlot(int)
    def fill_progress_bar_download(self, value):
        '''
        Заггрузка прогресс бара скачивания
        :param value: инкремент
        :return:
        '''
        prev = self.progressBar_download.value()
        self.progressBar_download.setValue(prev + value)

    @QtCore.pyqtSlot(int)
    def set_progress_bar_max(self, maxvalue):
        '''
        Задание максимального значения прогресс бару
        :param maxvalue: максимальное значение
        :return:
        '''
        self.progressBar.setMaximum(maxvalue)

    @QtCore.pyqtSlot(int)
    def set_progress_bar_max_download(self, maxvalue):
        '''
        Задание максимального значения прогресс бару загрузки
        :param maxvalue: максиамальное значение
        :return:
        '''
        self.progressBar_download.setMaximum(maxvalue)

    @QtCore.pyqtSlot(bool)
    def set_access_settings_and_commands(self, flag):
        '''
        Управление областью параметров и меню баром
        :param flag: флаг состояния
        :return:
        '''
        self.param_area.setEnabled(flag)
        self.menuBar.setEnabled(flag)

    @QtCore.pyqtSlot(str)
    def set_status_bar(self, value):
        '''
        Загрузка сособщения в статус бар
        :param value: сообщение
        :return:
        '''
        self.statusBar.showMessage(value, -1)

    def can_choice_action(self):
        '''
        Метод отслеживания изменения состояния чекбоксов для выбора источника
        :return:
        '''
        if self.shallalist_check.checkState() or self.digincore_check.checkState():
            self.action_import.setEnabled(True)
            self.action_clean.setEnabled(True)
        else:
            self.action_import.setEnabled(False)
            self.action_clean.setEnabled(False)

    def refresh_information(self):
        '''
        Обновление информации
        :return:
        '''
        self.set_access_settings_and_commands(False)
        try:
            pythoncom.CoInitialize()
            object_tmg = win32com.client.Dispatch('FPC.Root')
            isa_array = object_tmg.GetContainingArray()
        except pythoncom.pywintypes.com_error as ce:
            self.statusBar.showMessage('Отсутствует подключение к Forefront TMG', -1)
            self.set_access_settings_and_commands(False)
            return
        else:
            rule = isa_array.ArrayPolicy.PolicyRules.Item(self.choice_ruleName.currentText())

            rule_sets = rule.AccessProperties.DestinationDomainNameSets

            counter = 1

            self.domainNameSetsTMG.clear()
            self.operations.clear()

            self.progressBar.setValue(0)
            self.progressBar_download.setValue(0)

            while (counter <= rule_sets.Count):
                domain_set_name = rule_sets.Item(counter).Name
                item = QtWidgets.QListWidgetItem(domain_set_name)
                self.domainNameSetsTMG.addItem(item)
                counter += 1

            self.set_access_settings_and_commands(True)

    def start_refresh_information(self):
        threading.Thread(target=self.refresh_information).start()

    def start_thread_clean(self):
        threading.Thread(target=self.clean_tmg).start()

    def start_thread_import(self):
        self.progressBar.setValue(0)
        self.progressBar_download.setValue(0)

        threading.Thread(target=self.import_tmg).start()

    def clean_tmg(self):
        '''
        Удаление наборов доменных имён по префиксам
        :return:
        '''
        self.set_access_settings_and_commands(False)
        mini_dict = {'shallalist_check': self.SHL_prefix.text(),
                     'digincore_check': self.DGNC_prefix.text()}

        for i in [self.shallalist_check, self.digincore_check]:
            if i.checkState():
                processSources.Blacklist(mini_dict[i.objectName()], self.choice_ruleName.currentText(),
                                         emit_dict=self.emit_dict).clean()

        self.set_access_settings_and_commands(True)

    def import_tmg(self):
        '''
        Обноление наборов
        :return:
        '''
        mini_dict = {'shallalist_check': self.SHL_prefix.text(),
                     'digincore_check': self.DGNC_prefix.text()}

        for i in [self.shallalist_check, self.digincore_check]:
            if i.checkState():
                self.set_access_settings_and_commands(False)
                blacklist = processSources.Blacklist(mini_dict[i.objectName()], self.choice_ruleName.currentText(),
                                                     emit_dict=self.emit_dict)
                bl = mpro.Process(blacklist.import_tmg(i.objectName().replace('_check', ''),
                                                       ['COPYRIGHT', 'global_usage'],
                                                       exceptions_domains=(self.exceptionsDomains.toPlainText().split(' ')),
                                                       part_size=self.part_size.value()))
                bl.daemon = True
                bl.start()
