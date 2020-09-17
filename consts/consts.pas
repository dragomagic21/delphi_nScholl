
{*******************************************************}
{                                                       }
{       Borland Delphi Visual Component Library         }
{                                                       }
{  Copyright (c) 1995-2001 Borland Software Corporation }
{                                                       }
{*******************************************************}

unit Consts;

interface

resourcestring
  SOpenFileTitle = 'Открыть';
  SCantWriteResourceStreamError = 'Can''t write to a read-only resource stream';
  SDuplicateReference = 'WriteObject called twice for the same instance';
  SClassMismatch = 'Ресурс %s - неверный класс';
  SInvalidTabIndex = 'Индекс Tab''а выходит за границы';
  SInvalidTabPosition = 'Позиция Tab''а несовместима с текущем стилем Tab''а ';
  SInvalidTabStyle = 'Стиль Tab''a несовместим с текущей позицией Tab''a';
  SInvalidBitmap = 'Неверное растровое(Bitmap) изображение ';
  SInvalidIcon = 'Неверное изображение иконки(Icon)';
  SInvalidMetafile = 'Неверное изображение метафайла';
  SInvalidPixelFormat = 'Неверный формат пикселей (PixelFormat)';
  SInvalidImage = 'Неверное изображение';
  SBitmapEmpty = 'Пустое изображение';
  SScanLine = 'Индекс сканирования линий вне границ';
  SChangeIconSize = 'Невозможно изменение размеров иконки';
  SOleGraphic = 'Неверная операция к TOleGraphic';
  SUnknownExtension = 'Неизвестный формат изображения (.%s)';
  SUnknownClipboardFormat = 'Неподдерживаемый буфером формат';
  SOutOfResources = 'Вне ресурсов системы';
  SNoCanvasHandle = 'Канва (Canvas) не позволяет рисовать';
  SInvalidImageSize = 'Неверный размер изображения';
  STooManyImages = 'Слишком много изображений';
  SDimsDoNotMatch = 'Размеры компонента TImage не соответствуют размерам TImageList';
  SInvalidImageList = 'Неверный ImageList';
  SReplaceImage = 'Недоступна смена изображения';
  SImageIndexError = 'Неверный индекс компонента ImageList';
  SImageReadFail = 'Ошибка чтения данных ImageList из потока';
  SImageWriteFail = 'Ошибка записи данных ImageList в поток';
  SWindowDCError = 'Ошибка создания контекста окна';
  SClientNotSet = 'Клиент TDrag не инициализирован';
  SWindowClass = 'Ошибка создания класса окна';
  SWindowCreate = 'Ошибка создания окна';
  SCannotFocus = 'Невозможно естановить фокус ввода на отключённом или невидимом окне';
  SParentRequired = 'Control ''%s'' не имеет родительской формы';
  SParentGivenNotAParent = 'Переданный родитель не родитель ''%s''';
  SMDIChildNotVisible = 'Невозможно спрятать дочернюю форму MDI (многодокументного приложения)';
  SVisibleChanged = 'Свойство Visible нельзя изменять в процедурах OnShow или OnHide';
  SCannotShowModal = 'Модальное окно нельзя сделать видимым (Visible:=true)';
  SScrollBarRange = 'Свойство ScrollBar''а вне допустимых границ';
  SPropertyOutOfRange = '%s свойство вне допустимых границ';
  SMenuIndexError = 'Индекс меню вне пределах границ';
  SMenuReinserted = 'Меню дважды вставлено';
  SMenuNotFound = 'Под-меню не найдено';
  SNoTimers = 'Не хватает доступных таймеров';
  SNotPrinting = 'Принтер сейчас не находится в режиме печати';
  SPrinting = 'Идёт печать';
  SPrinterIndexError = 'Индекс принтера вне допустимых границ';
  SInvalidPrinter = 'Возможно выделенный принтер не есть принтер';
  SDeviceOnPort = '%s на %s';
  SGroupIndexTooLow = 'GroupIndex не может быть меньше чем предыдущий объект меню GroupIndex';
  STwoMDIForms = 'В одном приложении не может быть более чем одна MDI форма';
  SNoMDIForm = 'Не могу создать форму. Ни одна из MDI форм не активирована';
  SImageCanvasNeedsBitmap = 'Изображение можно изменить только если оно является растровым(Bitmap) ';
  SControlParentSetToSelf = 'Объект не может быть родителем сам себе';
  SOKButton = 'OK';
  SCancelButton = 'Отмена';
  SYesButton = '&Да';
  SNoButton = '&Нет';
  SHelpButton = '&Помощь';
  SCloseButton = '&Закрыть';
  SIgnoreButton = '&Игнорировать';
  SRetryButton = '&Повтор';
  SAbortButton = 'Прекратить';
  SAllButton = '&Все';

  SCannotDragForm = 'Не могу перемещать форму';
  SPutObjectError = 'Перенесите объект к неустановленному предмету';
  SCardDLLNotLoaded = 'Не могу загрузить CARDS.DLL';
  SDuplicateCardId = 'Найдено двойное CardId ';

  SDdeErr = 'Ошибка от DDE  ($0%x)';
  SDdeConvErr = 'DDE ошибка - переговоры не установлены ($0%x)';
  SDdeMemErr = 'Ошибка вызванная когда оперативная память DDE превысила допустимую ($0%x)';
  SDdeNoConnect = 'Невозможно связаться с DDE переговорами';

  SFB = 'FB';
  SFG = 'FG';
  SBG = 'BG';
  SOldTShape = 'Невозможно загрузить раннюю версию TShape';
  SVMetafiles = 'Метафайлы';
  SVEnhMetafiles = 'Enhanced Metafiles';
  SVIcons = 'Иконки';
  SVBitmaps = 'Растр';
  SGridTooLarge = 'Таблица слишком большая для операций';
  STooManyDeleted = 'Удалено слишком много колонок или столбцов';
  SIndexOutOfRange = 'Индекс таблицы вне добустимых границ';
  SFixedColTooBig = 'Количество фиксированных колонок должно быть меньше, чем самих колонок';
  SFixedRowTooBig = 'Количество фиксированных столбцов должно быть меньше, чем самих столбцов';
  SInvalidStringGridOp = 'Невозможно вставить или удалить стобец из таблицы';
  SInvalidEnumValue = 'Неверная Enum величина';
  SInvalidNumber = 'Неверная нумерация';
  SOutlineIndexError = 'Индекс вне линий не найден';
  SOutlineExpandError = 'Родитель должен быть расширен';
  SInvalidCurrentItem = 'Неверная величина для данного объекта';
  SMaskErr = 'Неверная вводная величина';
  SMaskEditErr = 'Неверная вводная величина. Используйте Escape, чтобы отменить изменения';
  SOutlineError = 'Invalid outline index';
  SOutlineBadLevel = 'Неверное присвоение уровня';
  SOutlineSelection = 'Неверное выделение';
  SOutlineFileLoad = 'Ошибка загрузки файла';
  SOutlineLongLine = 'Линия слишком длинная';
  SOutlineMaxLevels = 'Достигнута максимальная глубина';

  SMsgDlgWarning = 'Внимание';
  SMsgDlgError = 'Ошибка';
  SMsgDlgInformation = 'Информация';
  SMsgDlgConfirm = 'Подтверждение';
  SMsgDlgYes = '&Да';
  SMsgDlgNo = '&Нет';
  SMsgDlgOK = 'OK';
  SMsgDlgCancel = 'Отмена';
  SMsgDlgHelp = '&Помощь';
  SMsgDlgHelpNone = 'Помощь недоступна';
  SMsgDlgHelpHelp = 'Помощь';
  SMsgDlgAbort = '&Отменить';
  SMsgDlgRetry = '&Повтор';
  SMsgDlgIgnore = '&Игнорировать';
  SMsgDlgAll = '&Все';
  SMsgDlgNoToAll = 'Не применять ко всем';
  SMsgDlgYesToAll = 'Применить ко всем';

  SmkcBkSp = 'BkSp';
  SmkcTab = 'Tab';
  SmkcEsc = 'Esc';
  SmkcEnter = 'Enter';
  SmkcSpace = 'Space';
  SmkcPgUp = 'PgUp';
  SmkcPgDn = 'PgDn';
  SmkcEnd = 'End';
  SmkcHome = 'Home';
  SmkcLeft = 'Left';
  SmkcUp = 'Up';
  SmkcRight = 'Right';
  SmkcDown = 'Down';
  SmkcIns = 'Ins';
  SmkcDel = 'Del';
  SmkcShift = 'Shift+';
  SmkcCtrl = 'Ctrl+';
  SmkcAlt = 'Alt+';

  srUnknown = '(Unknown)';
  srNone = '(None)';
  SOutOfRange = 'Величина должна быть между %d и %d';

  SDateEncodeError = 'Неверный аргумент для декодировки даты';
  SDefaultFilter = 'Все файлы (*.*)|*.*';
  sAllFilter = 'Все';
  SNoVolumeLabel = ': [ - no volume label - ]';
  SInsertLineError = 'Невозможно вставить линию';

  SConfirmCreateDir = 'Установленная директория не существует. Создать?';
  SSelectDirCap = 'Выберите директорию';
  SDirNameCap = '&Имя директории:';
  SDrivesCap = 'Ди&ски:';
  SDirsCap = '&Директории:';
  SFilesCap = '&Файлы: (*.*)';
  SNetworkCap = 'Се&ть...';

  SColorPrefix = 'Color';               //!! obsolete - delete in 5.0
  SColorTags = 'ABCDEFGHIJKLMNOP';      //!! obsolete - delete in 5.0

  SInvalidClipFmt = 'Неверный формат буфера';
  SIconToClipboard = 'Буфер не поддерживает изображение иконок (Icons)';
  SCannotOpenClipboard = 'Невозможно открыть буфер';

  SDefault = 'По умолчанию';

  SInvalidMemoSize = 'Текст превысил емкость компонента Memo';
  SCustomColors = 'Обычные цвета';
  SInvalidPrinterOp = 'Не поддерживаемая операция к выделенному принтеру';
  SNoDefaultPrinter = 'Здесь не выбран принтер';

  SIniFileWriteError = 'Невозможна запись к %s';

  SBitsIndexError = 'Индекс битов вне допустимых границ';

  SUntitled = '(Untitled)';

  SInvalidRegType = 'Неверный тип данных для ''%s''';

  SUnknownConversion = 'Неизвестный файл RichEdit (.%s)';
  SDuplicateMenus = 'Меню ''%s'' уже используется другой формой';

  SPictureLabel = 'Изображение:';
  SPictureDesc = ' (%dx%d)';
  SPreviewLabel = 'Предыдущее';

  SCannotOpenAVI = 'Невозможно открыть AVI';

  SNotOpenErr = 'Нет открытых MCI механизмов';
  SMPOpenFilter = 'Все файлы (*.*)|*.*|Wave файлы (*.wav)|*.wav|Midi файлы (*.mid)|*.mid|Видео для Windows (*.avi)|*.avi';
  SMCINil = '';
  SMCIAVIVideo = 'AVIVideo';
  SMCICDAudio = 'CDAudio';
  SMCIDAT = 'DAT';
  SMCIDigitalVideo = 'DigitalVideo';
  SMCIMMMovie = 'MMMovie';
  SMCIOther = 'Другое';
  SMCIOverlay = 'Аналоговое устройство';
  SMCIScanner = 'Сканер';
  SMCISequencer = 'MIDI-файл';
  SMCIVCR = 'Видеокассета';
  SMCIVideodisc = 'Виде-компакт-диск';
  SMCIWaveAudio = 'Wav-файл';
  SMCIUnknownError = 'Нейзвестная ошибка';

  SBoldItalicFont = 'Bold Italic';
  SBoldFont = 'Bold';
  SItalicFont = 'Italic';
  SRegularFont = 'Regular';

  SPropertiesVerb = 'Свойство';

  SServiceFailed = 'Service failed on %s: %s';
  SExecute = 'execute';
  SStart = 'start';
  SStop = 'stop';
  SPause = 'pause';
  SContinue = 'continue';
  SInterrogate = 'interrogate';
  SShutdown = 'shutdown';
  SCustomError = 'Service failed in custom message(%d): %s';
  SServiceInstallOK = 'Service installed successfully';
  SServiceInstallFailed = 'Service "%s" failed to install with error: "%s"';
  SServiceUninstallOK = 'Service uninstalled successfully';
  SServiceUninstallFailed = 'Service "%s" failed to uninstall with error: "%s"';

  SInvalidActionRegistration = 'Invalid action registration';
  SInvalidActionUnregistration = 'Invalid action unregistration';
  SInvalidActionEnumeration = 'Invalid action enumeration';
  SInvalidActionCreation = 'Invalid action creation';

  SDockedCtlNeedsName = 'Docked control must have a name';
  SDockTreeRemoveError = 'Error removing control from dock tree';
  SDockZoneNotFound = ' - Dock zone not found';
  SDockZoneHasNoCtl = ' - Dock zone has no control';

  SAllCommands = 'All Commands';

  SDuplicateItem = 'List does not allow duplicates ($0%x)';

  STextNotFound = 'Текст не найден: "%s"';
  SBrowserExecError = 'No default browser is specified';

  SColorBoxCustomCaption = 'Custom...';

  SMultiSelectRequired = 'Multiselect mode must be on for this feature';

  SKeyCaption = 'Ключ';
  SValueCaption = 'Величина';
  SKeyConflict = 'Ключ с именем "%s" Уже существует';
  SKeyNotFound = 'Ключ "%s" не найден';
  SNoColumnMoving = 'goColMoving is not a supported option';
  SNoEqualsInKey = 'Key may not contain equals sign ("=")';

  SSendError = 'Ошибка отправки сообщения';
  SAssignSubItemError = 'Cannot assign a subitem to an actionbar when one of it''s parent''s is already assigned to an actionbar';
  SDeleteItemWithSubItems = 'Item %s has subitems, delete anyway?';
  SDeleteNotAllowed = 'You are not allowed to delete this item';
  SMoveNotAllowed = 'Item %s is not allowed to be moved';
  SMoreButtons = 'Больше кнопок';
  SErrorDownloadingURL = 'Ошибка загрузки URL: %s';
  SUrlMonDllMissing = 'Невозможно загрузить %s';
  SAllActions = '(All Actions)';
  SNoCategory = '(No Category)';
  SExpand = 'Expand';
  SErrorSettingPath = 'Ошибка установки пути: "%s"';
  SLBPutError = 'Attempting to put items into a virtual style listbox';
  SErrorLoadingFile = 'Error loading previously saved settings file: %s'#13'Would you like to delete it?';
  SResetUsageData = 'Reset all usage data?';
  SFileRunDialogTitle = 'Run';
  SNoName = '(No Name)';
  SErrorActionManagerNotAssigned = 'ActionManager must first be assigned';
  SAddRemoveButtons = '&Add or Remove Buttons';
  SResetActionToolBar = 'Reset Toolbar';
  SCustomize = '&Customize';
  SSeparator = 'Separator';
  SCirularReferencesNotAllowed = 'Circular references not allowed';
  SCannotHideActionBand = '%s does not allow hiding';
  SErrorSettingCount = 'Error setting %s.Count';
  SListBoxMustBeVirtual = 'Listbox (%s) style must be virtual in order to set Count';
  SUnableToSaveSettings = 'Unable to save settings';
  SRestoreDefaultSchedule = 'Would you like to reset to the default Priority Schedule?';
  SNoGetItemEventHandler = 'No OnGetItem event handler assigned';
  SInvalidColorMap = 'Invalid Colormap this ActionBand requires ColorMaps of type TCustomActionBarColorMapEx';
  SDuplicateActionBarStyleName = 'Стиль с именем %s уже зарегистрирован';
  SStandardStyleActionBars = 'Стандартный стиль';
  SXPStyleActionBars = 'XP Стиль';
  SActionBarStyleMissing = 'No ActionBand style unit present in the uses clause.'#13 +
    'Your application must include either XPStyleActnCtrls, StdStyleActnCtrls or ' +
    'a third party ActionBand style unit in its uses clause.';

implementation

end.
