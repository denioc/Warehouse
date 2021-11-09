# pip install kivymd==0.104.2
# pip install kivy (1.11.1; 2.0.0)
# pip install openpyxl

# Настройка конфигурации
#from kivy.config import Config
#print(Config.get('graphics', 'position'))
#print(Config.get('graphics', 'resizable'))
#Config.set('graphics', 'position', 'auto') # default 'auto'
#Config.set('graphics', 'resizable', '1') # default '1'

# Импорт модуля графического интерфейса
from GUI import TopToolbar, Table, MenuMain, CalendarDialog, AppSettings
from GUI import icon_setup1, icon_setup2
# Импорт модуля работы с файлами Excel
from func_tools import *
from AppSettings import window_settings, files_settings
#from work_xls import Excel
# Импорт модулей kivy и kivymd
from kivymd.app import MDApp
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.navigationdrawer import MDNavigationLayout
from kivymd.uix.label import MDLabel
#from kivymd.uix.button import MDButton
from kivy.uix.screenmanager import ScreenManager, Screen
#from kivy.uix.anchorlayout import AnchorLayout
from kivy.core.window import Window
from kivy.utils import platform


class WarhouseApp(MDApp):
	"""
	+ Склад:
		- Приход
		- Расход
		- План (?)
	+ Закупка:
		- Подсчет
		- Анализ
	+ Себестоимость:
		- Расчет по закупке
		- Анализ по ценам

	==== Добавить обработку развернутого окна.
	"""
	use_kivy_settings = False	# Отключает отображение конфигурации kivy в настройках
	window_maximize = False		# Отображает, развернуто ли окно во весь экран

	def build_config(self, config):
		#if self._is_desktop():
		#	fscrn = 'False'
		#else: fscrn = 'auto'
		config.setdefaults('window', {
			'maximize': False,
			#'fullscreen': fscrn,
			'width': '1200',
			'height': '800',
			'left': '400',
			'top': '100'
		})
		config.setdefaults('files', {
			'main_xls': u'Общий.xlsx'
		})

	def build_settings(self, settings):
		settings.add_json_panel('Настройки окна', self.config, data=window_settings, icon=icon_setup1)
		settings.add_json_panel('Настройки файлов', self.config, data=files_settings, icon=icon_setup2)
		#self.settings = settings.add_json_panel('Настройки окна', self.config, data=window_settings, icon=icon_setup1)
		#self.settings = settings.add_json_panel('Настройки файлов', self.config, data=files_settings, icon=icon_setup2)

	def build(self):
		self._desktop = self._is_desktop()
		# Переопределяем функции при событии развернуть окно во весь экран и восстановить обратно
		Window.on_maximize = self.on_maximize
		Window.on_restore = self.on_restore
		# Настройка приложения
		config = self.config
		#config.read('warhouse.ini')
		self.window_maximize = config.getboolean('window', 'maximize')
		if self._desktop:
			if self.window_maximize:
				Window.maximize()
			else:
				self.set_window_conf()

		# Настраиваем панель настроек
		self.settings_cls = AppSettings
		#self.settings_cls = SettingsWindow # Настройки с левым меню

		# Установить цветовую тему
		#self.theme_cls.primary_palette = 'Purple' #'Green', 'Red', 'Blue'
		#self.theme_cls.theme_style = 'Light' # 'Dark'
		# Создаем объекты главного экрана
		screen_navigation = MDNavigationLayout()  #(FloatLayout) для управления расположением разных экранов
		screen_manager = ScreenManager()        #Диспетчер экранов. Для управления несколькими экранами
		screen_Main = Screen()	# Главное окно программы

		bl = MDBoxLayout(orientation='vertical')
		MainToolBar = TopToolbar(self)	# Шапка экрана
		self.StatusLabel = MDLabel(text="MBLabel", color = (1,0,0,1), halign = 'center')
		self.StatusLabel.size_hint = (1, 0.1)
		self.table = Table()

		#TestButton = MDButton(text="MDButton", on_press=self.press_test_button)

		self.Menu = MenuMain(self)
		#self.Settings = MenuSettings(self)
		#self.Settings = self.settings_cls

		# Добавляем объекты главного экрана
		screen_navigation.add_widget(screen_manager)
		screen_manager.add_widget(screen_Main)
		screen_Main.add_widget(bl)
		bl.add_widget(MainToolBar)
		#bl.add_widget(self.Workspace)
		bl.add_widget(self.StatusLabel)
		bl.add_widget(self.table)
		#bl.add_widget(TestButton)
		screen_navigation.add_widget(self.Menu)
		#screen_navigation.add_widget(self.Settings)

		# Создаем прочие объекты
		self.Calendar = CalendarDialog()
		self.Calendar.bind(on_save=self.on_close_Calendar)
		#self.xls = Excel()  #Добавить указание файла
		#self.xl = Excel

		# Чтение таблицы и загрузка данных в главное окно

		#print(f"MAIN: {self.Menu}")
		#print(f"MAIN: {dir(self.Menu.set_state)}")
		return screen_navigation

	def press_Menu(self, instance):
		print("MainMenu pressed")
		self.Menu.set_state('open')

	def press_Menu_1(self, instance):
		print(f"Pressed Menu1")
		wb = Excel.open_book(u'Общий.xlsx')
		ws = wb['Итог'] #wb['SPR4-03']
		#podschet()
		self.StatusLabel.text = "Pressed Menu1"
		Excel.view_sheet(ws, self.table)
		self.Menu.set_state('close')

	def press_Menu_2(self, instance):
		print(f"Pressed Menu2")
		self.StatusLabel.text = "Pressed Menu2"
		self.Menu.set_state('close')

	def press_Menu_3(self, instance):
		print(f"Pressed Menu3")
		self.StatusLabel.text = "Pressed Menu3"
		#self.open_settings()
		self.Menu.set_state('close')

	def press_Settings(self, instance):
		print("Settings pressed")
		self.open_settings()
		#self.Settings.set_state('open')

	# def press_Setup_1(self, instance):
	# 	print(f"Pressed Setup1")
	# 	self.StatusLabel.text = "Pressed Setup1"
	# 	self.open_settings()

	# def press_Setup_2(self, instance):
	# 	print(f"Pressed Setup2")
	# 	self.StatusLabel.text = "Pressed Setup2"

	# def press_Setup_3(self, instance):
	# 	print(f"Pressed Setup3")
	# 	self.StatusLabel.text = "Pressed Setup3"

	def press_Calendar(self, instance):
		print("Calendar pressed")
		#print(f"CAL: {self}")
		#print(f"CAL: {dir(self)}")
		self.Calendar.open()

	def on_close_Calendar(self, instance, value, date_range):
		print("Calendar closed:")
		print(f"    Selected Day: {self.Calendar.sel_day}")
		print(f"    Selected Month: {self.Calendar.sel_month}")
		print(f"    Selected Year: {self.Calendar.sel_year}")
		print(f"    Value: {value}")
		print(f"    Date_range: {date_range}")
		#print(f"CAL: {dir(self.Calendar)}")

	def on_maximize(self):
		# Функция вызывается по событию при разворачивании окна во весь экран
		self.window_maximize = not self.window_maximize

	def on_restore(self):
		# Функция вызывается по событию при восстановлении окна из развернутого или свернутого состояния
		self.window_maximize = False
		self.set_window_conf()
		
	def set_window_conf(self):
		# Устанавливает параметры размера и позиции экрана
		ww = self.config.getint('window', 'width')
		wh = self.config.getint('window', 'height')
		Window.size = (ww, wh)
		Window.left = self.config.getint('window', 'left')
		Window.top = self.config.getint('window', 'top')

	def on_stop(self):
		if self._desktop and not self.window_maximize:
			ww, wh = Window.size
			self.config.setall('window', {
				'width': ww,
				'height': wh,
				'left': Window.left,
				'top': Window.top
				})
		self.config.setall('window', {
			'maximize': self.window_maximize,
			#'fullscreen': Window.fullscreen,
			})
		self.config.write()

	def _is_desktop(self):
		# Функция проверяет, запущено ли приложение на ПК
		if platform in ('linux', 'win', 'macosx'):
			return True
		else: return False
	

		
	"""
	def press_test_button(self, instance):
		print("TestButton pressed")
		wb = self.xl.open_book(u'Общий.xlsx')
		ws = self.xl.read_sheet(wb,  'SPR4-03')
		head_color = (0.8,0.8,0.8,1)	# Цвет фона заголовка таблицы
		#table_color = (0.8,0.98,0.98,1)	# Цвет фона основной таблицы
		table_color = (1.0,1.0,1.0,1.0)
		# Настраиваем таблицу
		self.table.columns_width = {0: 100, 1: 100, 2:300, 3:350, 4:150, 5:150}	#индивидуальная ширина столбцов
		self.table.headers = [{'text': f"head {x}", 'background_color': head_color, 'focus': False} for x in range(6)]
		self.table.data = [{'text': f"data {x}", 'background_color': table_color, 'focus': False} for x in range(102)]
	"""



if __name__ == '__main__':
	WarhouseApp().run()

# Запуск приложения с другими настройками экрана
#python main.py --size=720x1294 --dpi=320 -w