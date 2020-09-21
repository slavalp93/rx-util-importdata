# rx-util-importdata
Репозиторий с шаблоном разработки «Утилита импорта данных».

## Описание
Шаблон позволяет перенести исторические данные в Directum RX. 
В качестве источника данных для переноса используются книги Excel с расширением XLSX.
Чтобы произвести миграцию документов и справочников в Directum RX из заменяемой системы, достаточно заполнить специально сформированные шаблоны Excel и запустить утилиту из командной строки с необходимыми параметрами.

### На текущий момент реализована возможность импорта следующих типов документов:
1. Договоры.
2. Дополнительные соглашения.
3. Входящие письма.
4. Исходящие письма.
5. Приказы.

### Справочников:
1. Организации (Контрагенты).
2. Наши организации.
3. Подразделения.
4. Должности.
5. Сотрудники.
6. Персоны.

## Модификация

Для модификации утилиты требуется наличие на рабочем месте разработчика:
1. Visual Studio 2017 и выше.
2. Установленый пакет .net framework 4.6.2 и выше.

Модификация выполняется за счет наследования и доработки реализованных классов:
* класс Entity - базовый абстрактный класс, от которого наследованы остальные классы;
* классы справочников (Databooks):  BusinessUnit, Company, Department, Employee, Person - реализуют процесс импорта исторических данных в справочники системы Directum RX;
* классы документов (EDocs): Contract, IncomingLetter, Order, OutgoingLetter, SupAgreement - реализуют процесс импорта документов с телами (или без тел) в систему Directum RX;
* методы, которые прямо не относятся к классам сущностей реализуются в классе BusinessLogic;
* механизмы работы с XLSX реализованы в классе ExcelProcessor. В качестве механизма работы с XLSX используется библиотека OpenXml.
* общие механизмы работы с сущностями реализованы в классах EntityProcessor и EntityWrapper.

## Подготовка инсталлятора

Для сборки инсталлятора используется утилита NSIS (https://github.com/kichik/nsis).
В папке решения https://github.com/DirectumCompany/rx-util-importdata/tree/master/Install находится готовый конфигурационный файл DirRxInstaller.nsi для генерации инсталлятора.

## Порядок установки и использования

Порядок установки и использования описан в документе https://github.com/DirectumCompany/rx-util-importdata/blob/master/doc/Instructions_for_loading_data_into_DirectumRX_rus.pdf

## Рекомендации по доработке утилиты (для партнеров)

Поскольку доступ к текущему репозиторию ограничен, для доработки утилиты под нужды проекта рекомендуется:
1. Создать собственный репозиторий для кастомизируемой утилиты импорта (на github или не другом подобном ресурсе).
2. Выполнить клонирование репозитория из https://github.com/DirectumCompany/rx-util-importdata в репозиторий кастомизируемой утилиты.
3. После клонирования репозитория сделать новую ветку из ветки соответствующей версии. При выходе новой версии Directum RX, формируется новая ветка для утилиты соответтсвующей версии https://github.com/DirectumCompany/rx-util-importdata/branches.
4. Выполнить кастомизацию и тестирование утилиты, после чего сформировать запрос на вытягивание в ветку нужной версии.
5. Выполнить рецензирование запроса на вытягивание с ответственным специалитом внутри вашей компании.
6. Сформировать готовый пакет с кастомизированной утилитой.
