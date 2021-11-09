"""
Модуль описания графического интерфейса приложения
"""
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.gridlayout import MDGridLayout
from kivymd.uix.toolbar import MDToolbar
from kivymd.uix.datatables import MDDataTable
from kivymd.uix.navigationdrawer import MDNavigationDrawer
from kivymd.uix.list import MDList, OneLineListItem, OneLineIconListItem, IconLeftWidget #, IconLeftWidgetWithoutTouch
from kivymd.uix.picker import MDDatePicker
from kivymd.uix.label import MDLabel
from kivymd.uix.boxlayout import MDBoxLayout
#from kivymd.uix.button import MDRoundFlatButton, MDFillRoundFlatButton
from kivy.properties import NumericProperty
from kivy.uix.settings import Settings, ContentPanel
from kivy.graphics import Color, Rectangle
from kivy.properties import ObjectProperty
from kivy.uix.boxlayout import BoxLayout

from kivy.uix.image import Image
from kivy.uix.scrollview import ScrollView
from kivy.uix.anchorlayout import AnchorLayout

from kivy.metrics import dp     #Времмено, возмоджно потом надо будет удалить

from table import WorkTable

"""
#Проверка иконки на наличие среди стандартных иконок kivyMD
from kivymd.icon_definitions import md_icons    #список иконок
find_icon = "content-save-settings"
if find_icon in md_icons.keys():
	print(f"Иконка {find_icon} есть в списке!")
else: 
	print(f"Иконки {find_icon} нет в списке:")
	# Выводим список с похожими именами иконок. Поиск ведется по первым 2м буквам заданной иконки (:3)
	print(list(filter(lambda x: x.startswith(find_icon[:3]), md_icons.keys())))
"""

# Определяем глобальные переменные
# Задаем иконки. Для указания стандартной иконки входящей в kivyMD достаточно указать ее название.
icon_logo = "data/logo/kivy-icon-256.png"
icon_menu = 'menu'
icon_settings = 'cogs'
icon_calendar = 'calendar-clock'
icon_setup1 = 'tune'
icon_setup2 = 'content-save-settings'
icon_setup3 = 'database-settings'

class TopToolbar(MDToolbar):
	"""Создает шапку с кнопками в верхней части экрана"""
	def __init__(self, root, **kwargs):
		self.root = root
		self.title = "TopToolbar"
		self.anchor_title = 'center'
		self.left_action_items = [[icon_menu, self.root.press_Menu]]
		self.right_action_items = [[icon_calendar, self.root.press_Calendar], 
			[icon_settings, self.root.press_Settings]]
		super().__init__(**kwargs)

class Table(WorkTable):
	""" Описание и настройка таблицы"""
	def __init__(self, **kwargs):
		super().__init__(**kwargs)
		#self.bg_color = (0, 0.5, 0.3, 1)

class TableView(MDDataTable):
	"""Описание виджета вывода таблицы kivymd. 
	Таблица работает как всплывающее окно, поэтому нам не подходит."""
	def __init__(self, **kwargs):
		#super().__init__(**kwargs)
		#self.size_hint=(0.5, 0.5)
		self.size_hint = (None, None)
		self.size = (600, 200)
		self.auto_dismiss = True
		self.column_data=[
			("Column 1", dp(20)),
			("Column 2", dp(30)),
			("Column 3", dp(50)),
			("Column 4", dp(100)),
			]
		self.row_data=[[1,1,1,1],[2,2,2,2],[3,3,3,3], [4,4,4,4]]
		super().__init__(**kwargs)

class MenuItem(OneLineListItem):
	"""Описание виджета для отображения отдельного пункта в меню"""
	def __init__(self, **kwargs):
		super().__init__(**kwargs)
		self.bg_color = (0.53, 0.71, 0.85, 1)    #Фон виджета

class MenuMain(MDNavigationDrawer):
	"""Создает главное меню команд, всплывающее в левой части экрана"""
	def __init__(self, root, **kwargs):
		self.root = root
		super().__init__(**kwargs)
		self.width = 200
		self.md_bg_color = (0.59, 0.91, 0.92, 0.3) #self.root.theme_cls.primary_light
		box = MDBoxLayout(orientation='vertical')
		box.padding = "4dp"
		box.spacing = "8dp"
		# Создание списка пунктов меню
		sv = ScrollView(do_scroll_y = True)
		list_items = MDList()
		# Добавляем пункты меню
		list_items.add_widget(MenuItem(text="Menu 1", on_press=self.root.press_Menu_1))
		list_items.add_widget(MenuItem(text="Menu 2", on_press=self.root.press_Menu_2))
		list_items.add_widget(MenuItem(text="Menu 3", on_press=self.root.press_Menu_3))
		# Добавляем объекты в меню
		sv.add_widget(list_items)
		box.add_widget(sv)
		self.add_widget(box)
		#print(f"MENU: {self.theme_cls.theme_style}")
		#print(f"MENU: {dir(self)}")

# class MenuSettings(MDNavigationDrawer):
# 	"""Создает меню настроек приложения, всплывающее в правой части экрана"""
# 	def __init__(self, root, **kwargs):
# 		self.root = root
# 		super().__init__(**kwargs)
# 		self.anchor = 'right'
# 		self.width = 400
# 		self.md_bg_color = (0.59, 0.91, 0.92, 0.3) #self.root.theme_cls.primary_light
# 		box = MDBoxLayout(orientation='vertical')
# 		box.padding = "8dp"
# 		box.spacing = "8dp"
# 		# Создаем объекты для заголовка с аватаром
# 		al = AnchorLayout()
# 		al.anchor_x = 'left'
# 		al.size_hint_y = None
# 		al.height = 100 # avatar.height
# 		img = Image(source='mylogo.png')
# 		img.size_hint = (None, None)
# 		img.size = ("64dp", "64dp")
# 		img.source = icon_logo #"kivymd.png"
# 		al.add_widget(img)
# 		user_name = MDLabel(text = "User Name")
# 		user_name.font_style = "Button"
# 		user_name.size_hint_y = None
# 		user_name.height = "10dp" #user_name.texture_size[1]
# 		user_name_note = MDLabel(text = "User Name Note")
# 		user_name_note.font_style = "Caption"
# 		user_name_note.size_hint_y = None
# 		user_name_note.height = "8dp" #user_name_note.texture_size[1]
# 		# Создание списка пунктов меню
# 		sv = ScrollView(do_scroll_y = True)
# 		list_items = MDList()
# 		# Добавляем пункты меню
# 		list_items.add_widget(SetupItem(icon=icon_setup1, text="Настройка окна", on_press=self.root.press_Setup_1))
# 		list_items.add_widget(SetupItem(icon=icon_setup2, text="Setup 2 dfsgfdgg", on_press=self.root.press_Setup_2))
# 		list_items.add_widget(SetupItem(icon=icon_setup3, text="Setup 3", on_press=self.root.press_Setup_3))
# 		# Добавляем объекты в меню
# 		sv.add_widget(list_items)
# 		box.add_widget(al)
# 		box.add_widget(user_name)
# 		box.add_widget(user_name_note)
# 		box.add_widget(sv)
# 		self.add_widget(box)
# 		#print(f"MENU: {self.theme_cls.theme_style}")
# 		#print(f"MENU: {dir(self)}")

class CalendarDialog(MDDatePicker):
	"""Клас диалогового окна календаря"""
	def __init__(self, **kwargs):
		super().__init__(**kwargs)
"""
class SettingsWindow(SettingsWithSidebar):
	#Класс окна настроек
	def __init__(self, **kwargs):
		super().__init__(**kwargs)
		with self.canvas.before:
			Color(0.59, 0.91, 0.92, 0.3)
			self.stp_rect = Rectangle(size=self.size, pos=self.pos)
		self.bind(size=self._upd_stp_rect, pos=self._upd_stp_rect)
		print(dir(self))
		print(self.size, self.pos)

	def _upd_stp_rect(self, instance, value):
		self.stp_rect.size = self.size
		self.stp_rect.pos = self.pos
"""

# class SettingMenu(MDBoxLayout):
# 	"""Создает" меню настроек приложения, всплывающее в правой части экрана"""
# 	selected_uid = NumericProperty(0)
# 	'''The uid of the currently selected panel.'''
# 	#close_button = ObjectProperty(MDRoundFlatButton)
# 	'''(internal) Reference to the widget's Close button.'''
# 	def __init__(self, root, **kwargs):
# 		self.root = root
# 		self.orientation = 'vertical'
# 		#self.anchor = 'right'
# 		self.width = 300
# 		self.md_bg_color = (0.59, 0.91, 0.92, 0.7) #self.root.theme_cls.primary_light
# 		self.padding = "8dp"
# 		self.spacing = "8dp"
# 		super().__init__(**kwargs)
# 		# Создаем объекты для заголовка с аватаром
# 		al = AnchorLayout()
# 		al.anchor_x = 'left'
# 		al.size_hint_y = None
# 		al.height = '65dp' #200 # avatar.height
# 		img = Image(source='mylogo.png')
# 		img.size_hint = (None, None)
# 		img.size = ("64dp", "64dp")
# 		img.source = icon_logo #"kivymd.png"
# 		al.add_widget(img)
# 		user_name = MDLabel(text = "User Name")
# 		user_name.font_style = 'Button'
# 		user_name.size_hint_y = None
# 		user_name.height = '10dp' #user_name.texture_size[1]
# 		user_name_note = MDLabel(text = "User Name Note")
# 		user_name_note.font_style = "Caption"
# 		user_name_note.size_hint_y = None
# 		user_name_note.height = '8dp' #user_name_note.texture_size[1]
# 		# Создание списка пунктов меню
# 		sv = ScrollView(do_scroll_y = True)
# 		self.list_items = MDList()
# 		sv.add_widget(self.list_items)
# 		# Кнопка закрыть настройки
# 		#self.close_button = MDRoundFlatButton(text='Close')
# 		self.close_button = MDFillRoundFlatButton(text='Close Settings')
# 		#self.close_button.font_size = '10dp'
# 		self.close_button.pos_hint = {'center_x': 0.5}
# 		# Добавляем объекты в меню
# 		self.add_widget(al)
# 		self.add_widget(user_name)
# 		self.add_widget(user_name_note)
# 		self.add_widget(sv)
# 		self.add_widget(self.close_button)
		
# 	def add_item(self, name, uid, icon=None):
# 		self.list_items.add_widget(SetupItem(icon=icon, text=name, uid=uid, on_press=self.select_panel))
		
# 	def select_panel(self, instance):
# 		self.selected_uid = instance.uid

from kivymd.uix.backdrop import MDBackdrop
#kivymd.uix.backdrop.backdrop.MDBackdropToolbar(**kwargs)
from kivymd.uix.backdrop import MDBackdropFrontLayer
from kivymd.uix.backdrop import MDBackdropBackLayer
#from kivymd.uix.chip import MDChip, MDChooseChip
#from kivymd.uix.button import MDRoundFlatIconButton
from kivy.animation import Animation
from typing import NoReturn
from kivy.properties import StringProperty
from kivy.clock import Clock
#from kivy.animation import Animation
import json
from kivy.config import ConfigParser
from kivy.uix.settings import SettingsPanel #as SetPan
from kivy.uix.settings import SettingItem
from kivy.uix.settings import SettingTitle, SettingNumeric
from kivy.compat import string_types, text_type
from kivymd.uix.textfield import MDTextField

class mySettingNumeric(SettingNumeric):
	def __init__(self, **kwargs):
		super().__init__(**kwargs)
		print(dir(self.children[0].children[1]))
		print(self.children[0].children[0].children)
		self.children[0].children[1].color = (1,0,0,1)
		self.children[0].children[0].children[0].color = (0,1,0,1)

class mySettingTitle(SettingTitle):
	def __init__(self, **kwargs):
		super().__init__(**kwargs)
		self.color = (0,0,0,1) # Работает
		#print(dir(self.title))
		#print(self.title)
		#print(self.color)
		
class SettingItemNumeric(SettingItem, MDTextField):
	def __init__(self, **kwargs):
		self.use_bubble = False
		super().__init__(**kwargs)
		self.hint_text = self.title
		self.helper_text = self.desc
		self.text = self.value
		self.helper_text_mode = 'on_focus' #"persistent"
		self.mode = 'fill' #'rectangle'
		self.font_size = '14sp'
		#self.hint_text_color_normal = (1,0,0,1) # Не имеет смысла
		self.current_hint_text_color = (0,0,0,1) # !только до kivymd v1.0.0
		#self.color_mode = 'accent' #'primary', ‘accent’, ‘custom’
		self.bind(focus=self._validate)

	def _validate(self, instance, value):
		# Функция сохраняет измененные значения когда снимается фокус
		if value:
			return
		is_float = '.' in str(self.value)
		try:
			if is_float:
				self.value = text_type(float(self.text))
			else:
				self.value = text_type(int(self.text))
		except ValueError:
			return


class AppSettings(Settings):
	def __init__(self, *args, **kwargs):
		self.interface_cls = SettingInterface
		super().__init__(*args, **kwargs)
		#self._types = {}
		#self.add_interface()
		#self.register_type('string', SettingString)
		#self.register_type('bool', SettingBoolean)
		self.register_type('numeric', SettingItemNumeric)
		#self.register_type('options', SettingOptions)
		#self.register_type('title', SettingTitle)
		#self.register_type('path', SettingPath)
		#self.register_type('color', SettingColor)
		#self._types[tp] = cls
		self.register_type('title', mySettingTitle)

		self.bind(parent=self.on_open)
		#print(dir(self))
		
	def add_json_panel(self, title, config, filename=None, data=None, icon=None):
		'''Переопределенная функция из Settings. 
		Для добавления иконки к кнопкам меню'''
		panel = self.create_json_panel(title, config, filename, data)
		uid = panel.uid
		if self.interface is not None:
			self.interface.add_panel(panel, title, uid, icon)

	# def create_json_panel(self, title, config, filename=None, data=None):
	# 	'''Create new :class:`SettingsPanel`.
	# 	.. versionadded:: 1.5.0
	# 	Check the documentation of :meth:`add_json_panel` for more information.
	# 	'''
	# 	if filename is None and data is None:
	# 		raise Exception('You must specify either the filename or data')
	# 	if filename is not None:
	# 		with open(filename, 'r') as fd:
	# 			data = json.loads(fd.read())
	# 	else:
	# 		data = json.loads(data)
	# 	if type(data) != list:
	# 		raise ValueError('The first element must be a list')
	# 	panel = SettingsPanel(title=title, settings=self, config=config)
	# 	for setting in data:
	# 		# determine the type and the class to use
	# 		if 'type' not in setting:
	# 			raise ValueError('One setting are missing the "type" element')
	# 		ttype = setting['type']
	# 		cls = self._types.get(ttype)
	# 		if cls is None:
	# 			raise ValueError(
	# 				'No class registered to handle the <%s> type' %
	# 				setting['type'])
	# 		# create a instance of the class, without the type attribute
	# 		del setting['type']
	# 		str_settings = {}
	# 		for key, item in setting.items():
	# 			str_settings[str(key)] = item
	# 		print(str_settings)
	# 		instance = cls(panel=panel, **str_settings)
	# 		# instance created, add to the panel
	# 		panel.add_widget(instance)
	# 		#print(dir(instance))
	# 		#print(instance, instance.children)
	# 		#print(instance, ':')
	# 		#for obj in instance.walk():
	# 		#	print(obj)
	# 		#print('---')
	# 	return panel
		
	def on_close(self):
		pass
		
	def on_open(self, *args):
		if self.parent:
			Clock.schedule_once(lambda x: self.children[0].dispatch('on_load'))

class SettingInterface(MDBackdrop):
	#__events__ = ('on_close', )
	closing_time = NumericProperty(0.2)
	closing_transition = StringProperty("out_quad")
	
	def __init__(self, **kwargs):
		super().__init__(**kwargs)
		self.title = 'Настройки'
		self.left_action_items = [['menu', lambda x: self.open()]] #Работет только первая кнопка
		self.right_action_items = [[icon_setup3, lambda x: print("Pressed icon 3")], 
			[icon_setup2, lambda x: self.parent.dispatch('on_close')]]
		back_layer = MDBackdropBackLayer()
		front_layer = MDBackdropFrontLayer()
		#print(dir(front_layer))
		#print(front_layer)
		#with front_layer.canvas:
		#	Color(1,0,0,1)
		#front_layer.opacity = 0
		self.menu = SettingMenu(self)
		#self.menu.size_hint = (0.5, 1)
		#self.menu.width = 300
		self.content = ContentPanel()
		#self.content = SettingContent()
		back_layer.add_widget(self.menu)
		front_layer.add_widget(self.content)
		#self.bind(size=front_layer.setter('size'))
		self.add_widget(back_layer)
		self.add_widget(front_layer)
		self.menu.bind(selected_uid=self.content.setter('current_uid'))
		self.register_event_type('on_load')
		
	def add_panel(self, panel, name, uid, icon=None):
		self.menu.add_item(name, uid, icon=icon)
		#self.menu.add_widget(SettingMenuItem(icon=icon, text=name, uid=uid)) #, on_press=self.select_panel))
		self.content.add_panel(panel, name, uid)
		
	def on_load(self):
		""" Открывает слой back (меню) при загрузке панели настроек """
		self.animtion_icon_menu()	# on kivymd==0.104.2
		y = self.ids.header_button.height - self.ids._front_layer.height
		Animation(y=y, d=0, t='linear').start(self.ids._front_layer)
		self._front_layer_open = True

	def close(self) -> NoReturn:
		"""Переопределнная функция. Открывает слой front"""
		Animation(y=0, d=self.closing_time, t=self.closing_transition).start(
			self.ids._front_layer)
		self._front_layer_open = False
		#self.dispatch("on_close")

	def on_close(self, *args):
		pass

	def on_open(self, *args):
		pass
		
class SettingMenu(MDBoxLayout):
	"""Создает меню настроек приложения, всплывающее в правой части экрана"""
	selected_uid = NumericProperty(0)
	'''The uid of the currently selected panel.'''
	#close_button = ObjectProperty(MDRoundFlatButton)
	def __init__(self, root, **kwargs):
		self.root = root
		self.orientation = 'vertical'
		#self.md_bg_color = (0.59, 0.91, 0.92, 0.7) #self.root.theme_cls.primary_light
		self.padding = "8dp"
		self.spacing = "8dp"
		super().__init__(**kwargs)
		# Создаем объекты для заголовка с аватаром
		al = AnchorLayout()
		al.anchor_x = 'left'
		al.size_hint_y = None
		al.height = '65dp' # avatar.height
		img = Image(source='mylogo.png')
		img.size_hint = (None, None)
		img.size = ("64dp", "64dp")
		img.source = icon_logo #"kivymd.png"
		al.add_widget(img)
		user_name = MDLabel(text = "User Name")
		user_name.font_style = "Button"
		user_name.size_hint_y = None
		user_name.height = "10dp" #user_name.texture_size[1]
		user_name_note = MDLabel(text = "User Name Note")
		user_name_note.font_style = "Caption"
		user_name_note.size_hint_y = None
		user_name_note.height = "8dp" #user_name_note.texture_size[1]
		# Создание списка пунктов меню
		sv = ScrollView(do_scroll_y = True)
		self.list_items = MDList()
		sv.add_widget(self.list_items)
		# Добавляем объекты в меню
		self.add_widget(al)
		self.add_widget(user_name)
		self.add_widget(user_name_note)
		self.add_widget(sv)
		
	def add_item(self, name, uid, icon=None):
		self.list_items.add_widget(SettingMenuItem(
			icon=icon, 
			text=name, 
			uid=uid, 
			on_press=self.select_panel
			))
		
	def select_panel(self, instance):
		self.selected_uid = instance.uid
		self.root.open()
		
class SettingMenuItem(OneLineIconListItem):
	"""Создание виджета для отображения отдельного пункта в меню"""
	uid = NumericProperty(0)
	def __init__(self, icon=None, **kwargs):
		iwid = IconLeftWidget()
		#iwid = IconLeftWidgetWithoutTouch() # in version 1.0.0
		#iwid.size = (50, 50)
		#iwid.padding = (100,100, 100, 100)
		#iwid.anchor_x = 'left'
		#iwid.children[0].halign = 'left'
		if icon:
			iwid.icon = icon
		super().__init__(**kwargs)
		self.bg_color = (0.53, 0.71, 0.85, 1)    #Фон виджета
		#			self.size_hint = (None, None)
		self.add_widget(iwid)
		#print(self.size, self.size_hint)
		#print(iwid.children[0].pos, iwid.children[0].halign)
		#print("padding:", iwid.children[0].padding)
		#print(iwid.__events__)
		#iwid.children[0].bind(padding = lambda s, v: print("iw_padd", v))
		#iwid.children[0].bind(pos = lambda s, v: print("iw_pos", v))
		#print(self.pos_hint, self.pos)
		#print(iwid.anchor_x)
		#print(dir(self))


# class SettingsPanel(MDGridLayout):
# 	''' Переоппределенный класс от SettingsPanel. Установлен MDGridLayout и сделан черный фон.
# 	По хорошему, надо изменить цвет текста на черный и убрать этот переопределенный класс.'''

# 	title = StringProperty('Default title')
# 	'''Title of the panel. The title will be reused by the :class:`Settings` in
# 	the sidebar.
# 	'''

# 	config = ObjectProperty(None, allownone=True)
# 	'''A :class:`kivy.config.ConfigParser` instance. See module documentation
# 	for more information.
# 	'''

# 	settings = ObjectProperty(None)
# 	'''A :class:`Settings` instance that will be used to fire the
# 	`on_config_change` event.
# 	'''

# 	def __init__(self, **kwargs):
# 		kwargs.setdefault('cols', 1)
# 		#self.md_bg_color = (0,0,0,1)	# Черный цвет фона
# 		#self.specific_text_color = (1,0,0,1) # Не помогает
# 		#self.specific_secondary_text_color = (1,0,0,0.7) # Не помогает
# 		super(SettingsPanel, self).__init__(**kwargs)
# 		#print(dir(self))
# 		#print(self.specific_text_color)
# 		#print(self.specific_secondary_text_color)

# 	def on_config(self, instance, value):
# 		if value is None:
# 			return
# 		if not isinstance(value, ConfigParser):
# 			raise Exception('Invalid config object, you must use a'
# 							'kivy.config.ConfigParser, not another one !')

# 	def get_value(self, section, key):
# 		'''Return the value of the section/key from the :attr:`config`
# 		ConfigParser instance. This function is used by :class:`SettingItem` to
# 		get the value for a given section/key.
# 		If you don't want to use a ConfigParser instance, you might want to
# 		override this function.
# 		'''
# 		config = self.config
# 		if not config:
# 			return
# 		return config.get(section, key)

# 	def set_value(self, section, key, value):
# 		current = self.get_value(section, key)
# 		if current == value:
# 			return
# 		config = self.config
# 		if config:
# 			config.set(section, key, value)
# 			config.write()
# 		settings = self.settings
# 		if settings:
# 			settings.dispatch('on_config_change',
# 							config, section, key, value)

# class SettingContent(ContentPanel):
# 	def __init__(self, **kwargs):
# 		super().__init__(**kwargs)
# 		#print(self.size_hint, self.pos_hint)
# 		#print(dir(self.container))
# 		#print(self.container)

# class SettingMenuItem(MDChip):
# 	"""Создание виджета для отображения отдельного пункта в меню"""
# 	uid = NumericProperty(0)
# 	def __init__(self, icon=None, text='', **kwargs):
# 		self.text = text
# 		self.icon = icon
# 		super().__init__(**kwargs)
# 		self.bg_color = (0.53, 0.71, 0.85, 1)    #Фон виджета
# 		self.selected_chip_color = (1, 0, 0, 1)


