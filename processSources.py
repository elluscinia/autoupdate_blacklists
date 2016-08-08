'''
Модуль с для обработки блэклистов из различных источников
'''
import io
import multiprocessing as mpro
import os
import sys
import time
from queue import Empty

import requests
import win32com.client

import handler
from makeLog import logger


def mock(value):
    '''
    Декоратор для замещения возвращаемых значений
    позволяет не вызывать функцию, а сделать вид что она сработала и вернула указанное значение
    В случае отсутсвия доступа в Интернет возвращает ошибку и останавливает выполнение скрипта
    '''
    def mocker(func):
        def mc(*args, **kwargs):
            return open(value, 'rb')
        try:
            answer = requests.get('http://google.com', timeout=1)
        except:
            logger.info('Ошибка подключения. Похоже интернета нет. Используем скачанные файлы')
            return mc
        else:
            if answer.status_code == 407:
                return mc
            else:
                return func
    return mocker


def download(source, emit_dict=None):
    '''
    Cкачивание чёрных листов из выбранного источника
    :param source: назание источника
    :param emit_dict: словарь для обращения к функциям GUI
    :return:
    '''
    if emit_dict:
        emit_dict['operation_GUI'].emit('Скачивание из %s' % source)

    if source == 'shallalist':
        logger.info('Скачивание из shallalist')
        url = 'http://www.shallalist.de/Downloads/shallalist.tar.gz'
        try:
            response = requests.get(url, stream=True)
        except Exception as ex:
            logger.info('Ошибка подключения')
            if emit_dict:
                emit_dict['operation_GUI'].emit('Ошибка подключения')
            sys.exit()
        else:
            if response.status_code == 407:
                logger.info('Ошибка подключения к сети')
                if emit_dict:
                    emit_dict['operation_GUI'].emit('Ошибка подключения')
                sys.exit()

    elif source == 'digincore':
        logger.info('Скачивание из digincore')
        url = 'http://www.digincore.org/index.php/zagruzit-bleklisty/file/13-squid-porn'
        data = {'license_agree': 'on',
                'submit': '%D0%97%D0%B0%D0%B3%D1%80%D1%83%D0%B7%D0%B8%D1%82%D1%8C',
                'download': '13'}

        from lxml import html
        with requests.Session() as s:
            page = s.get(url)
            p = html.fromstring(page.content)
            form = p.xpath('//form')[-1]
            key = form.xpath('input')[3].items()[1][1]
            data[key] = '1'
            s.headers.update({'Referer': url})

            try:
                response = s.post(url, data=data, stream=True)
            except Exception as ex:
                logger.info('Ошибка подключения')
                if emit_dict:
                    emit_dict['operation_GUI'].emit('Ошибка подключения')
                return
            else:
                if response.status_code == 407:
                    logger.info('Ошибка подключения')
                    if emit_dict:
                        emit_dict['operation_GUI'].emit('Ошибка подключения')
                    return

    total_length = response.headers.get('content-length')
    if emit_dict:
        emit_dict['progressBar_max_download'].emit(int(total_length))
    start = time.clock()
    dl = 0
    total_length = int(total_length)
    bytestream = io.BytesIO()
    for chunk in response.iter_content(1024):
        dl += len(chunk)
        bytestream.write(chunk)
        done = int(50 * dl / int(total_length))
        if emit_dict:
            emit_dict['fill_progressBar_download'].emit(len(chunk))
        sys.stdout.write("\r[%s%s] %s bps" % ('=' * done, ' ' * (50-done), dl//(time.clock() - start)))
        sys.stdout.flush()
    print('\n')
    bytestream.seek(0)
    return bytestream


def unpacker(bin_data, exception_list, source, emit_dict=None):
    '''
    Распаковка архивов в память
    :param bin_data:
    :param exception_list: лист исключений
    :param extraction: расширение файла
    :return: генератор
    '''
    if source == 'digincore':
        import zipfile
        try:
            archive_ref = zipfile.ZipFile(bin_data)
        except zipfile.BadZipfile as err:
            logger.info('BAD ZIP FILE')
            if emit_dict:
                emit_dict['operation_GUI'].emit('BAD ZIP FILE')
            raise StopIteration
        else:
            for member in archive_ref.namelist():
                if archive_ref.getinfo(member).compress_type and (os.path.basename(member) not in exception_list):
                    yield archive_ref.open(member), member

    if source == 'shallalist':
        import tarfile
        try:
            archive_ref = tarfile.open(fileobj=bin_data)
        except tarfile.ReadError as err:
            logger.info('BAD TAR FILE: %s', err)
            if emit_dict:
                emit_dict['operation_GUI'].emit('BAD TAR FILE')
            raise StopIteration
        else:
            for member in archive_ref.getmembers():
                if member.isfile() and (os.path.basename(member.name) not in exception_list):
                    yield archive_ref.extractfile(member), member.name


class Blacklist():
    '''
    Класс загрузки/обработки чёрных списков с 2 сайтов: http://www.digincore.org/ & http://www.shallalist.de
    :param prefix: префикс
    :param ruleName: имя правила
    :param emit_dict: словарик для выбора сигнала в GUI
    '''

    def __init__(self, prefix, rule_name, emit_dict=None):
        self.prefix = prefix
        self.rule_name = rule_name
        self.emit_dict = emit_dict

    def listener(self, callback, information):
        '''
        Вызывает методы GUI
        :param callback: название метода из списка
        :param information: передаваемый методу параметр
        :return:
        '''
        if self.emit_dict:
            self.emit_dict[callback].emit(information)

    def update_progress(self, result_tuple):
        '''
        Метод для обработки значения, возвращаемого очередью
        :param result_tuple: кортеж значений
        :return:
        '''
        key, val = result_tuple
        self.listener(key, val)

    def import_tmg(self, source, exception_list, exceptions_domains='', part_size=500000):
        '''
        Распаковывает каждый файл в ОП и парсит его
        :param source: источник импорта
        :param exception_list: лист исключений
        :param exceptions_domains: доменные имена - исключения
        :param part_size: макимальное кол-во обрабатываемых в одном файле URL
        :return:
        '''

        self.listener('set_StatusBar', 'Загрузка архива из источника')

        try:
            response = download(source, self.emit_dict)
        except Exception:
            error = 'Ошибка загрузки, проверьте соединение с сетью!'
            logger.info(error)
            self.listener('operation_GUI', error)
            self.listener('set_StatusBar', 'Ошибка скачивания файла')
            sys.exit()
        else:
            bin_data = response

        try:
            f = unpacker(bin_data, exception_list, source, self.emit_dict)
        except Exception:
            logger.info('Ошибка загрузки, сервер не доступен или загружает неверную информацию, '
                        'проверьте соединение с сетью!')
            self.listener('operation_GUI', 'Ошибка загрузки, сервер не доступен или загружает неверную информацию, '
                                           'проверьте соединение с сетью!')
            self.listener('unblock_area', True)
        else:
            self.clean()

            logger.info('Начался парсинг %s чёрных списков', source)
            self.listener('operation_GUI', 'Начался парсинг %s чёрных списков' % source)
            self.listener('set_StatusBar', 'Начался парсинг %s чёрных списков' % source)

            count_files = 0
            pool = mpro.Pool(1)
            manager = mpro.Manager()
            queue = manager.Queue()

            jobs = []
            all_row = 0
            for unreadContent, member in f:
                domain_set = []
                rows_counter = 0
                chunk_counter = 0
                for row in unreadContent:
                    if row in exceptions_domains:
                        continue
                    domain_set.append(row.decode('ISO-8859-1'))
                    rows_counter += 1
                    all_row += 1
                    if rows_counter > part_size:
                        domain_set_name = '%s_%s_%d' % (self.prefix, member.replace('/', '_'), chunk_counter)
                        jobs.append((sorted(domain_set), domain_set_name, self.rule_name, queue))
                        logger.debug('Создание XML: %s', domain_set_name)
                        self.listener('operation_GUI', 'Создание XML: %s' % domain_set_name)
                        chunk_counter += 1
                        rows_counter = 0
                        domain_set.clear()
                if chunk_counter:
                    domain_set_name = '%s_%s_%d' % (self.prefix, member.replace('/', '_'), chunk_counter)
                    jobs.append((sorted(domain_set), domain_set_name, self.rule_name, queue))
                    logger.info('Файл %s разделён на %d частей', self.prefix + '_' + member.replace('/', '_'),
                                chunk_counter + 1)
                    self.listener('operation_GUI', 'Файл %s разделён на %d частей' %
                                  (self.prefix + '_' + member.replace('/', '_'), chunk_counter + 1))
                else:
                    domain_set_name = '%s_%s' % (self.prefix, member.replace('/', '_'))
                    jobs.append((sorted(domain_set), domain_set_name, self.rule_name, queue))
                    logger.debug('Создание XML: %s', domain_set_name)
                    self.listener('operation_GUI', 'Создание XML: %s' % domain_set_name)
                count_files += 1

            logger.info('Импорт %d файлов', count_files)

            self.listener('operation_GUI', 'Импорт %d файлов' % count_files)
            self.listener('set_StatusBar', 'Импорт файлов в TMG')
            self.listener('progressBar_max', all_row * 2)

            pool.imap_unordered(handler.create_and_load_xml, jobs)
            pool.close()
            tasks_done = 0
            i = 0
            while 1:
                try:
                    self.listener('operation_GUI', 'Обрабатываем объект в очереди')
                    logger.debug('Обрабатываем объект в очереди')
                    task_result = queue.get(False)
                    if task_result is None:
                        self.listener('operation_GUI', 'Задание выполнено')
                        logger.debug('Задание выполнено')
                        tasks_done += 1
                    else:
                        self.update_progress(task_result)
                except Empty:
                    self.listener('operation_GUI', 'Очередь пуста')
                    logger.info('Очередь пуста')
                if tasks_done == len(jobs):
                    self.listener('operation_GUI', 'Все задания выполнены')
                    logger.debug('Все задания выполнены')
                    break
                i += 1
            pool.join()
            self.listener('unblock_area', True)
            self.listener('set_StatusBar', 'Импорт наборов доменных имён завершён!')
            logger.info('Импорт наборов доменных имён завершён!')

    def clean(self):
        '''
        Удаляет из TMG ранее загруженные с этого сайта блэклисты
        '''
        import pythoncom
        pythoncom.CoInitialize()
        self.listener('set_StatusBar', 'Удаление наборов')
        try:
            object_xml = win32com.client.Dispatch('FPC.Root')
            isa_array = object_xml.GetContainingArray()
        except Exception:
            logger.info('Ошибка в подключении к TMG')
        else:
            n_domain_name_sets = isa_array.RuleElements.DomainNameSets
            logger.info('Удаление %s чёрных списков TMG', self.prefix)
            count_deleted_files = 0

            try:
                rule = isa_array.ArrayPolicy.PolicyRules.Item(self.rule_name)
            except Exception:
                logger.info('Правила с названием %s для набора доменных имён не существует', self.rule_name)
                self.listener('operation_GUI',
                              'Правила с названием %s для набора доменных имён не существует' % self.rule_name)
                sys.exit()
            else:
                rule_sets = rule.AccessProperties.DestinationDomainNameSets

                counter = 1

                while(counter <= rule_sets.Count):
                    domain_set_name = rule_sets.Item(counter).Name
                    if domain_set_name.startswith(self.prefix):
                        rule_sets.Remove(counter)
                        rule.Save()
                        n_domain_name_sets.Remove(domain_set_name)
                        n_domain_name_sets.Save()

                        self.listener('clean_GUI', domain_set_name)

                        logger.debug('Файл %s удалён из TMG', domain_set_name)
                        self.listener('operation_GUI', 'Файл %s удалён из TMG' % domain_set_name)
                        count_deleted_files += 1
                    else:
                        counter += 1

                logger.info('%d наборов доменных имён было удалёно из TMG', count_deleted_files)
                self.listener('operation_GUI',
                              ('%d наборов доменных имён было удалено из правила %s TMG' %
                               (count_deleted_files, self.rule_name)))

                counter = 0
                for item in n_domain_name_sets:
                    if item.Name.startswith(self.prefix):
                        self.listener('operation_GUI', '%s набор доменных имён был удален из TMG' % str(item.Name))
                        logger.debug('%s набор доменных имён был удален из TMG', item.Name)
                        n_domain_name_sets.Remove(item.Name)
                        counter += 1
                n_domain_name_sets.Save()

                self.listener('operation_GUI', '%s наборов доменных имён было удалено из наборов TMG' % counter)
                logger.info('%s наборов доменных имён было удалено из наборов TMG', counter)

                self.listener('set_StatusBar', 'Удаление наборов доменных имён завершено!')
                logger.info('Удаление наборов доменных имён завершено!')
