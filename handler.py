'''

Модуль ссоздания XML файлов и импорта в TMG

'''

from uuid import uuid4
from xml.sax.saxutils import unescape
import sys

import pythoncom
import win32com.client

from makeLog import logger

xmlStart = """<?xml version="1.0" encoding="UTF-8"?><fpc4:Root xmlns:fpc4="http://schemas.microsoft.com/isa/config-4" xmlns:dt="urn:schemas-microsoft-com:datatypes" StorageName="FPC" StorageType="0"><fpc4:Build dt:dt="string">7.0.7734.100</fpc4:Build><fpc4:Comment dt:dt="string"/><fpc4:Edition dt:dt="int">32</fpc4:Edition><fpc4:EnterpriseLevel dt:dt="int">2</fpc4:EnterpriseLevel><fpc4:ExportItemClassCLSID dt:dt="string">{61A8568E-53C1-4D6D-BBD8-4F7150EB3093}</fpc4:ExportItemClassCLSID><fpc4:ExportItemCompatibilityVersion dt:dt="int">2</fpc4:ExportItemCompatibilityVersion><fpc4:ExportItemScope dt:dt="int">0</fpc4:ExportItemScope><fpc4:ExportItemStorageName dt:dt="string">{%(expGUID)s}</fpc4:ExportItemStorageName><fpc4:IsaXmlVersion dt:dt="string">7.3</fpc4:IsaXmlVersion><fpc4:OptionalData dt:dt="int">12</fpc4:OptionalData><fpc4:Upgrade dt:dt="boolean">0</fpc4:Upgrade><fpc4:ConfigurationMode dt:dt="int">0</fpc4:ConfigurationMode><fpc4:Arrays StorageName="Arrays" StorageType="0"><fpc4:Array StorageName="{9DABC2DD-2B86-4200-B856-F755E7441696}" StorageType="0"><fpc4:AdminMajorVersion dt:dt="int">0</fpc4:AdminMajorVersion><fpc4:AdminMinorVersion dt:dt="int">0</fpc4:AdminMinorVersion><fpc4:Components dt:dt="int">-1</fpc4:Components><fpc4:DNSName dt:dt="string"/><fpc4:Name dt:dt="string"/><fpc4:Version dt:dt="string">0</fpc4:Version><fpc4:RuleElements StorageName="RuleElements" StorageType="0"><fpc4:DomainNameSets StorageName="DomainNameSets" StorageType="0"><fpc4:DomainNameSet StorageName="{%(expGUID)s}" StorageType="1"><fpc4:DomainNameStrings>"""

xmlEnd = """</fpc4:DomainNameStrings><fpc4:Name dt:dt="string">%s</fpc4:Name></fpc4:DomainNameSet></fpc4:DomainNameSets></fpc4:RuleElements></fpc4:Array></fpc4:Arrays></fpc4:Root>"""


def create_and_load_xml(data):
    '''
    Создание и загрузка в TMG XML файлов с наборами доменных имён
    :param data: параметры: доменные имена, название набора, имя правила, вызовы функций для GUI
    :return:
    '''
    domains, domain_set_name, rule_name, queue = data

    logger.debug('Делаем набор доменных имён %s', domain_set_name)
    try:
        queue.put(('operation_GUI', 'Делаем набор доменных имён %s' % domain_set_name))
    except Exception as ex:
        logger.debug(ex)
    else:

        queue.put(('unpack_GUI', domain_set_name))

        file_dom = (xmlStart % {'expGUID': str(uuid4()).upper()})
        for url in domains:
            try:
                file_dom += ('<fpc4:Str dt:dt="string">' + (url.replace('&', '&amp;')) + '</fpc4:Str>')
            except Exception:
                logger.info('Ошибка создания файла XML %s', domain_set_name)
                queue.put(('operation_GUI', 'Ошибка создания файла XML %s' % domain_set_name))
                return
        file_dom += (xmlEnd % unescape(domain_set_name))

        dom = win32com.client.Dispatch('Msxml2.DOMDocument.3.0')
        dom.async = False
        dom.loadXML(file_dom)

        try:
            pythoncom.CoInitialize()
            object_tmg = win32com.client.Dispatch('FPC.Root')
        except pythoncom.pywintypes.com_error as ce:
            logger.info('Отсутствует подключение к Forefront TMG', -1)
            sys.exit()
        else:
            isa_array = object_tmg.GetContainingArray()
            isa_array.RuleElements.DomainNameSets.Import(dom, 0)

            queue.put(('fill_progressBar', len(domains)))

            logger.debug('Файл %s импортирован в набор доменных имён TMG', domain_set_name)
            queue.put(('operation_GUI', 'Файл %s импортирован в набор доменных имён TMG' % domain_set_name))

            try:
                rule = isa_array.ArrayPolicy.PolicyRules.Item(rule_name)
            except Exception:
                logger.info('Правила с названием %s для набора доменных имён не существует', rule_name)
                queue.put(('operation_GUI', 'Правила с названием %s для набора доменных имён не существует' % rule_name))
                sys.exit()
            else:
                rule_sets = rule.AccessProperties.DestinationDomainNameSets

                try:
                    rule_sets.Add(domain_set_name, 0)
                except Exception:
                    logger.debug('Набор с таким названием уже привязан к этому правилу')
                    queue.put(('operation_GUI', 'Набор с таким названием уже привязан к этому правилу'))
                else:
                    rule.Save()
                    queue.put(('import_GUI', domain_set_name))

                    logger.debug('Файл %s добавлен в правило %s', domain_set_name, rule_name)
                    queue.put(('operation_GUI', 'Файл %s добавлен в правило %s' % (domain_set_name, rule_name)))

                    queue.put(('fill_progressBar', len(domains)))
            queue.put(None)
            return len(domains)
