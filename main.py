'''

Программа для обновления чёрных списков с ресурсов http://www.digincore.org/ и http://www.shallalist.de

'''

import argparse
import sys
from multiprocessing import freeze_support

from processSources import Blacklist


def create_parser():
    __version__ = '1.2.0.2-280716'
    parser = argparse.ArgumentParser(add_help=False, prog='main.py',
                                     description='Программа позволяет импортировать чёрные списки из двух ресурсов: '
                                                 'http://www.digincore.org/ & http://www.shallalist.de '
                                                 'А также распаковывать скачанные архивы, парсить распакованные файлы, '
                                                 'создавать XML файлы и загружать их в TMG и удалять их оттуда.',
                                     epilog='(c) Соловьева Е.А. 2016 Версия %(__version__)s' % locals())
    parent_group = parser.add_argument_group(title='Параметры')
    parent_group.add_argument('--prefix', '-pf', metavar='prefix',
                              help='Параметр, определяющий источник импорта:SHL - shallalist.de и DGNC - digincore.org',
                              required=True)
    parent_group.add_argument('--rule', '-r', metavar='rule',
                              help='Имя правила Forefront TMG', required=True)
    parent_group.add_argument('--help', '-h', action='help', help='Справка')

    subparsers = parser.add_subparsers(dest='command', title='Команды', description='')

    import_parser = subparsers.add_parser('import', add_help=False, help='Импорт в правило чёрных списков из источника')
    import_group = import_parser.add_argument_group(title='Параметры import')
    import_group.add_argument('--source', '-s', metavar='source', choices=['shallalist', 'digincore'], required=True,
                              help='Выбор ресурса в качестве источника импорта')
    import_group.add_argument('--maxSize', '-max', metavar='maxSize', type=int, default=500000,
                              help='Кол-во URL в одном наборе доменных имен')
    import_group.add_argument('--exceptions', '-exp', metavar='exceptions', nargs='*',
                              help='Названия файлов-исключений правила Forefront TMG')
    import_group.add_argument('--exceptionsDomains', '-expD', metavar='exceptionsDomains', default=[''], nargs='*',
                              help='Исключаемые доменные имена для правила в Forefront TMG')
    import_group.add_argument('--help', '-h', action='help', help='Справка import')

    clean_parser = subparsers.add_parser('clean', add_help=False,
                                         help='Удаление доменных имен, начинающихся на указанный префикс, из TMG')
    clean_group = clean_parser.add_argument_group(title='Параметры clean')
    clean_group.add_argument('--help', '-h', action='help', help='Справка clean')

    return parser


def main():
    if len(sys.argv) == 1:
        from PyQt5 import QtWidgets, QtGui, QtCore
        from resources import resource_path
        import handlerInterface

        app = QtWidgets.QApplication(sys.argv)

        splash_pix = QtGui.QPixmap(resource_path('logo.png'))

        splash = QtWidgets.QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)

        splash.setMask(splash_pix.mask())

        splash.show()

        window = handlerInterface.Application()

        window.show()
        splash.finish(window)
        sys.exit(app.exec_())
    else:
        parser = create_parser()
        namespace = parser.parse_args(sys.argv[1:])
        blacklist = Blacklist(namespace.prefix, namespace.rule)

        if namespace.command == 'import':
            blacklist.import_tmg(namespace.source, set(['COPYRIGHT', 'global_usage'] + namespace.exceptions if
                                                       namespace.exceptions is not None else
                                                       ['COPYRIGHT', 'global_usage']),
                                 set(namespace.exceptionsDomains + ['']), namespace.maxSize)
        elif namespace.command == 'clean':
            blacklist.clean()
        else:
            parser.print_help()


if __name__ == '__main__':
    freeze_support()
    main()
