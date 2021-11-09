"""
Виджет вывода таблицы с данными.
TODO:
	- Добавить возможность изменять высоту строк вручную - аналогично ручной настройке ширины строк
	- Переделать ячейки наименований строк и столбцов с textinput в label
"""
# pip install kivy==1.11.1
# На 10.09.2021 вышла версия kivy 2.0.0. В ней исправлен глюк с bubble, но появился другой глюк - в recycleview при создании gridlayout с одной строкой (количество столбцов равно количеству ячеек) выводится только первая ячейка. Этот глюк уже исправлен и будет выложен в версии 2.1.0. А пока можно скачать с репозитория и заменить файл recycleview, либо пользоваться версией 1.11.1.

from kivy.uix.boxlayout import BoxLayout
from kivy.uix.recycleview.views import RecycleDataViewBehavior
from kivy.uix.recycleview import RecycleView
from kivy.uix.recyclegridlayout import RecycleGridLayout
from kivy.effects.scroll import ScrollEffect
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.graphics import Color, Line, Rectangle
from kivy.core.window import Window
from kivy.event import EventDispatcher
from kivy.clock import Clock
from kivy.metrics import dp, pt
from kivy import resources
from kivy.base import EventLoop
from kivy.core.text import Text as TXT
from kivy.properties import BooleanProperty
from kivy.properties import ListProperty
from kivy.properties import ObjectProperty
from kivy.properties import DictProperty
from kivy.properties import NumericProperty
from kivy.properties import OptionProperty
from kivy.properties import StringProperty
from kivy.properties import VariableListProperty

from kivy import __version__ as kivy_version
print("Kivy Version:", kivy_version)

if kivy_version == '1.11.1':
	# В kivy version=1.11.1 есть баг с появлением bubble. Исправлено в версии 2.0.0
	# У TextInput на Android есть глюк - при вызове Bubble, он появляется и тут же исчезает. На других устройствах работает нормально. Решением пока является костыль - при инициализации TextInput вызвать _show_cut_copy_paste с атрибутами TextInput (это его функция для вызова и инициализации Bubble).
	from kivy.uix.textinput import Selector
	# Для устранения дефекта с появлением bubble при нажатии на кнопки выделения текста handle,
	# нужно изменить класс Select и функцию _handle_released в TextInput
	def selector_on_touch_up(instance,touch):
		if touch.grab_current is instance:
			instance.transform_touch(touch)
			super(Selector, instance).on_touch_up(touch)

	Selector.always_release = True # для удобства перемещения handle
	Selector.on_touch_up = selector_on_touch_up

if kivy_version == '2.0.0':
	# В kivy version=2.0.0 появился баг в recycleview (исправление в RecycleGridLayout) - при 
	# количестве колонок равной размеру данных (rows=1) отображается только первая ячейка.
	# Баг отсутсвовал в версии 1.11.1 и будет исправлен в версии 2.1.0
	import itertools 
	chain_from_iterable = itertools.chain.from_iterable
	def compute_visible_views_v2_0_0(self, data, viewport):
			if self._cols_pos is None:
				return []
			x, y, w, h = viewport
			right = x + w
			top = y + h
			at_idx = self.get_view_index_at
			tl, tr, bl, br = sorted((
				at_idx((x, y)),
				at_idx((right, y)),
				at_idx((x, top)),
				at_idx((right, top)),
			))

			n = len(data)
			if len({tl, tr, bl, br}) < 4:
				# visible area is one row/column
				return range(min(n, tl), min(n, br + 1))
			indices = []
			stride = len(self._cols) if self._fills_row_first else len(self._rows)
			if stride:
				x_slice = br - bl + 1
				indices = chain_from_iterable(
					range(min(s, n), min(n, s + x_slice))
					for s in range(tl, bl + 1, stride))
			return indices

	RecycleGridLayout.compute_visible_views = compute_visible_views_v2_0_0


COL_DEFAULT_WIDTH = '64dp'	#Ширина столбца по умолчанию
ROW_DEFAULT_HEIGHT = '20dp'	#Высота строк по умолчанию
ROW_NAME_WIDTH = '32dp'	#Ширина поля наименования строк
RES_REGION_WIDTH = 20	# Ширина области для изменения размера
RES_REGION_HEIGHT = 10	# Высота области для изменения размера
DEFAULT_HEAD_BG_COLOR = (0.8,0.8,0.8,1) # цвет ячеек заголовка таблицы по умолчанию
DEFAULT_BODY_BG_COLOR = (0.8,0.98,0.98,1) # цвет ячеек тела таблицы по умолчанию
RES_LINE_COLOR = (1,0,0,1)	# Цвет линии новой ширины столбца
RES_DASH_OFFSET = 1	# Отступ у пунктирной линии новой ширины столбца. 0 - линия сплошная
CELL_BORDER_COLOR = (0,0,0,1)	# Цвет контура ячеек
CELL_BORDER_WIDTH = 1.0			# Ширина линии контура ячеек
BAR_WIDTH = 10	# Ширина полосы прокрутки. По умолчанию 2
PADDING = [2, 2, 0, 0]	# Отступы до текста: [padding_left, padding_top, padding_right, padding_bottom]
LINE_SPACING = 5		# Отступы между строк
print("LINE_SPACING =", LINE_SPACING)
CHARCTER_WIDTH_PX = 7	# Ширина символа '0' в пикселях (для расчета отображения шрифта)
CHARCTER_HEIGHT_PX = 15	# Высота символа в пикселях (для расчета отображения шрифта)
DEFAULT_FONT_NAME = 'Roboto'	# Имя шрифта по-умолчанию
#FONT_SIZE = 15 #'11pt'	# Размер шрифта по-умолчанию

""" Класс для создания области изменения размера. """
class Region(Rectangle):
	def __init__(self, parent, **kwargs): #parent, **kwargs):
		super().__init__(**kwargs)
		self.parent = parent
		
	# Метод рассчитывает, попадает ли заданная точка координат в эту область (по аналогии collide_point)
	def is_collide(self, x, y):
		pos_start_x = self.pos[0]
		pos_start_y = self.pos[1]
		pos_end_x = self.pos[0] + self.size[0]
		pos_end_y = self.pos[1] + self.size[1]
		if (pos_start_x < x < pos_end_x) and ((pos_start_y < y < pos_end_y)):
			return True
		else: return False
		
	def to_widget(self, x, y, relative=False):
		return self.parent.to_widget(x, y, relative)

"""
Класс события. Отслеживает положение курсора мыши. Когда курсор располагается над виджетом, 
он меняет свое отображение. Когда курсор вне виджета, меняет свое отображение на стандартное. 
"""
class EventMouseCursor(EventDispatcher):
	objects = ListProperty([])		# Список хранит объекты, для которых нужно прослушивать курсор
	#inside = BooleanProperty(False)	# Текущее состояние курсора. True: курсор установлен в 'size_we'
	cursor = OptionProperty('arrow', options=['arrow', 'ibeam', 'wait', 'crosshair', 'wait_arrow', 
		'size_nwse', 'size_nesw', 'size_we', 'size_ns', 'size_all', 'no', 'hand'])	# Текущий вид курсора.
	set_cur = {'horizontal': 'size_we', 'vertical': 'size_ns'}
	def __init__(self, **kwargs):
		# Регистрируем событие
		self.register_event_type('on_mouse_cursor')
		# При возникновении нескольких одинаковых событий, достаточно один раз отправить команду 
		# на изменение курсора. Поэтому используем триггер планировщика. 
		# Он запускает событие один раз в пределах одного кадра.
		self.trigger_set = Clock.create_trigger(self.set_mouse_cursor)
		super().__init__(**kwargs)
		Window.fbind('mouse_pos', self.on_mouse_pos)	# Отслеживаем курсор мыши

	def add_listen_obj(self, obj):
		""" Добавляет объект в список прослушиваемых объектов. 
		Т.к. список содержит ссылки объектов, то при удалении объкта, он сам исчезает из списка"""
		if obj not in self.objects:
			self.objects.append(obj)

	def clear_listen_objects(self):
		""" Удаляет все объекты из списка """
		self.objects.clear()
	
	def on_mouse_pos(self, instance, pos):
		""" Функция отслеживает положение мыши и выдает соответствующие команды на изменение курсора"""
		# Проверяем, находится ли курсор хотя бы над одним из объектов в списке
		obj_collide = [obj.is_collide(*obj.to_widget(*pos)) for obj in self.objects]
		if any(obj_collide):
			if self.cursor == 'arrow':
				type_obj = self.objects[obj_collide.index(True)].parent.rv.location
				self.cursor = self.set_cur[type_obj]
				self.trigger_set()
		else:
			if self.cursor != 'arrow':
				self.cursor = 'arrow'
				self.trigger_set()
				
	def set_mouse_cursor(self, *args):
		Window.set_system_cursor(self.cursor)

	def on_mouse_cursor(self):
		pass

""" Событие изменения размера столбцов таблицы """
class EventResizeColRow(EventDispatcher):
	def __init__(self, **kwargs):
		self.register_event_type('on_resize_col')
		self.register_event_type('on_resize_row')
		super().__init__(**kwargs)
		
	def on_resize_col(self, touch):
		pass
		
	def on_resize_row(self, touch):
		pass

""" 
Виджет элемента заголовка каждого столбца таблицы
"""
class CellHead(RecycleDataViewBehavior, TextInput):
    # По-умолчанию многострочный режим включен (multiline = True). 
    # Высота всех ячеек в строке должна соответствовать
    # Чтобы поставить индивидуальную высоту ячейки, нужно size_hint_y=None; height={высота}
	index = None
	rv = None
	use_bubble = False
	readonly = BooleanProperty(True) # Режим только для чтения, редактирование заголовков невозможно
	allow_copy = BooleanProperty(False) # Запрещаем копирование текста
	padding = VariableListProperty(PADDING)
	halign = 'center'
	background_normal = StringProperty('')
	rect = None
	resize_rect = None

	def init_border(self):
		# Чтобы экран был белее, убираем background_normal и рисуем контур самостоятельно.
		with self.canvas.before:
			Color(*CELL_BORDER_COLOR)
			self.rect = Line(rectangle=(self.x, self.y, self.width, self.height), width=CELL_BORDER_WIDTH)

	def init_resize_region_width(self):
		""" Инициализирует область изменения размера ширины столбца. 
			Когда мышь наводится на эту область, меняется курсор """
		with self.canvas:
			self.resize_rect = Region(self, pos=(self.right-RES_REGION_WIDTH/2, self.pos[1]), size=(RES_REGION_WIDTH, self.height))
		
	def init_resize_region_height(self):
		""" Инициализирует область изменения размера высоты строки. 
			Когда мышь наводится на эту область, меняется курсор """
		with self.canvas:
			self.resize_rect = Region(self, pos=(self.x, self.y-RES_REGION_HEIGHT/2), size=(self.width, RES_REGION_HEIGHT))
		
	def refresh_view_attrs(self, rv, index, data):
		self.index = index
		self.rv = rv
		if not self.rect:
			self.init_border()
		if not self.resize_rect:
			if self.rv.location == 'horizontal':
				self.init_resize_region_width()
				rv.event_mouse.add_listen_obj(self.resize_rect)	# Добавляяем объект в список прослушиваемых у события мыши
			elif self.rv.location == 'vertical':
				self.init_resize_region_height()
				rv.event_mouse.add_listen_obj(self.resize_rect)
		return super().refresh_view_attrs(rv, index, data)

	def refresh_view_layout(self, rv, index, layout, viewport):
		# Обновляем rect
		x, y, = layout['pos']
		w,h = layout['size']
		self.rect.rectangle = (x, y, w, h)
		# Обновляем resize_rect
		if rv.location == 'horizontal':
			right = x + w
			self.resize_rect.pos = (right-RES_REGION_WIDTH/2, y)
			self.resize_rect.size = (RES_REGION_WIDTH, h)
		elif rv.location == 'vertical':
			self.resize_rect.pos = (x, y-RES_REGION_HEIGHT/2)
			self.resize_rect.size = (w, RES_REGION_HEIGHT)
		return super().refresh_view_layout(rv, index, layout, viewport)

	def on_minimum_height(self, instance, value):
		if not self.rv.auto_height:		# Если отключен режим автоподстройки высоты строк
			return
		if self.height == value:
			return
		# Если идет процесс обновления layout, то данные м.б. не верные, изменение высоты игнорируем
		if not self.rv.is_updating:
			self.rv.cell_min_height[self.index] = value

	def on_focus(self, instance, value):
		# Если идет процесс обновления layout, то фокус м.б. не верным, изменение фокуса игнорируем
		if not self.rv.is_updating:
			self.rv.data[self.index]['focus'] = value
		if kivy_version=='1.11.1' and value and not self._bubble:
			self._show_cut_copy_paste(self.pos, self, True)	#для нормального функционирования Bubble

	def on_text(self, instance, text):
		if self.parent and self.focus:
			self.rv.data[self.index]['text'] = text
		
	def on_touch_down(self, touch):
		# Отслеживаем нажатие на область изменения размера, а так же тройное нажатие на ячейку
		if self.resize_rect.is_collide(*touch.pos):
			self.focus = False		# Убираем фокус на поле ввода ячейки
			self.start_resize(touch)
		else:	#Чтобы вне процедуры изменения размера сама ячейка тоже нажималась
			if self.rv.event_mouse.cursor == 'arrow':
				return super().on_touch_down(touch)

	def on_touch_up(self, touch):
		# Если нажали и быстро отпустили в области изменения размера, то событие touch_up остается здесь.
		# У родителя оно не проходит, но линия нового размера уже отрисовывается. 
		# На этот случай очищаем Canvas здесь
		if hasattr(touch.ud, 'line'):
			self.rv.parent.canvas.after.clear()
		super().on_touch_up(touch)
		
	def start_resize(self, touch):
		""" Запускает процедуру изменения размера ячейки """
		if self.rv:
			touch.ud['index'] = self.index
			if self.rv.location == 'horizontal':
				touch.ud['pos_x'] = self.to_window(*self.pos)[0]
				self.rv.event_resize.dispatch('on_resize_col', touch)
			elif self.rv.location == 'vertical':
				touch.ud['pos_y'] = self.to_window(0, self.top)[1]
				self.rv.event_resize.dispatch('on_resize_row', touch)
	
	if kivy_version == '1.11.1':
		# Переопределенная функция TextInput. Для работы bubble при нажатии на handle
		
		def _handle_released(self, instance):
			sf, st = self._selection_from, self.selection_to
			if sf == st:
				return
			#self._update_selection()
			pos = (instance.right if instance is self._handle_left else instance.x, instance.top + self.line_height)
			pos = self.to_parent(*pos, relative=True)
			self._show_cut_copy_paste(pos, EventLoop.window)


"""
Класс отображения заголовков столбцов таблицы.
Для быстрого вывода большого количества данных используется RecycleView (он уже включает ScrollView)
Автоматическая подстройка размера высоты строки добавлена с возможностью переделки под несколько строк.
Незначительное изменение кода позволит добавить несколько строк в заголовок таблицы
"""
class HeadTable(RecycleView):
	# Расположение заголовка м.б.'horizontal'- горизонтальное (по-умолчанию), или 'vertical'- вертикальное.
	# Расположение необходимо указать при создании экземпляра класса
	cell_min_height = DictProperty({})
	is_updating = BooleanProperty(False)
	event_mouse = ObjectProperty(None)
	auto_height = False		# Режим автоматической подстройки высоты строк
	
	def __init__(self, *args, **kwargs):
		self.event_mouse = args[0]		# Событие наведения мыши на правую область ячейки
		self.event_resize = args[1]		# Событие изменения ширины столбца
		self.location = kwargs.pop('location', 'horizontal')	# Расположение заголовка
		self.effect_cls = ScrollEffect  # Отключить эффект прокрутки страницы за пределы границ контента
		self.bar_width = 0              # Толщина полосы прокрутки. По умолчанию 2
		self.scroll_type = ['bars', 'content']  # Тип прокрутки - по перемещиню контента и ползунка
		self.grid = RecycleGridLayout()
		self.grid.spacing = 0
		self.grid.size_hint = (None, None)
		self.grid.bind(minimum_size = self.grid.setter('size'))
		self.grid.default_size_hint = (1, 1)
		self.grid.row_default_height = ROW_DEFAULT_HEIGHT	# Высота строк по-умолчанию
		if self.location == 'vertical':
			self.do_scroll_y = True	# Включить только вертикальную прокрутку
			self.size_hint = (None, 1)
			self.grid.cols = 1
			self.grid.padding = [2, 0, 0, 2]
			self.width = ROW_NAME_WIDTH
			self.grid.col_default_width = ROW_NAME_WIDTH	#Ширина столбцов по умолчанию
		else:
			self.do_scroll_x = True         # Включить только горизонтальную прокрутку
			self.size_hint = (1, None)	# Высота поля для заголовков должна задаваться жестко
			self.grid.rows = 1
			self.grid.padding = [0, 2, 2, 0]
			self.height = ROW_DEFAULT_HEIGHT
			self.grid.col_default_width = COL_DEFAULT_WIDTH	#Ширина столбцов по умолчанию
		super().__init__(**kwargs)
		self.add_widget(self.grid)
		# Настраиваем RecycleView для отображения заголовков
		self.viewclass = CellHead   # Класс виджета для отображения заголовков столбцов
		if self.auto_height:
			# Создаем триггер обновления layout, чтобы все изменения обрабатывались за один раз
			self.init_trigger()

	def init_trigger(self, *largs):
		# Триггер обновления layout, чтобы все изменения обрабатывались за один раз
		self.trigger_upd_grid = Clock.create_trigger(self._upd_grid)
		self.bind(data = self._set_rows_height_default)

	def on_cell_min_height(self, instance, value):
		""" При изменении высоты ячейки запускает процесс пересчета и обновления layout """
		self.trigger_upd_grid()

	def _upd_grid(self, *largs):
		""" Пересчитывает и обновляет layout """
		cols = len(self.data)	#self.grid.cols	# Изменить если нужно несколько строк
		need_upd = False
		for row_num in self.grid.rows_minimum:
			row = [self.cell_min_height.get(row_num*cols+key) for key in range(cols)]
			max_height = max(row)
			if max_height != self.grid.rows_minimum[row_num]:
				self.grid.rows_minimum.update({row_num: max_height})
				need_upd = True
		if need_upd:
			self.event_mouse.clear_listen_objects()	# Удаляем объекты события (при обновлении создаются новые)
			self.refresh_from_layout()	# Обновляем layout
			self.is_updating = True		# Помечаем, что идет процесс обновления
			Clock.schedule_once(self._end_updating)

	def _end_updating(self, *largs):
		self.is_updating = False

	def _set_rows_height_default(self, instance, value):
		""" Устанавливает высоту строк по-умолчанию """
		length = len(value)
		if length > 0:
			row_count = self.grid.rows	#int(length/self.grid.cols)	# Изменить если нужно несколько строк
			self.cell_min_height = {i: 0 for i in range(length)}
			rows = {i: self.grid.row_default_height for i in range(row_count)}
			self.grid.rows_minimum = rows

""" 
Виджет каждой ячеки таблицы данных.
У TextInput на Android есть глюк - при вызове Bubble, он появляется и тут же исчезает. На других устройствах работает нормально. Решением пока является костыль - при инициализации TextInput вызвать _show_cut_copy_paste с атрибутами TextInput (это его функция для вызова и инициализации Bubble).
"""
class CellData(RecycleDataViewBehavior, TextInput):
	# По-умолчанию многострочный режим включен (multiline = True).
	# При изменении высоты поля текста отправляем команду на обработку и изменение высоты строки, 
	# обновляем layout в следующем кадре.
	# Когда вводим текст вручную, то после обновления layout фокус может перейти на другую ячейку, 
	# и введенные данные потеряются. Поэтому необходимо перед обновлением сохранить фокус и перезаписать данные.
	index = None
	rv = None
	use_bubble = True
	padding = VariableListProperty(PADDING)	# Отступы по краям до текста
	line_spacing = NumericProperty(LINE_SPACING)		# Отступы между строк
	background_normal = StringProperty('')
	rect = None
	
	def init_border(self):
		# Чтобы экран был белее, убираем background_normal и отрисовывать контур самомтоятельно.
		with self.canvas.before:
			Color(*CELL_BORDER_COLOR)
			self.rect = Line(rectangle=(self.x, self.y, self.width, self.height), width=CELL_BORDER_WIDTH)

	def refresh_view_attrs(self, rv, index, data):
		self.index = index
		self.rv = rv
		if not self.rect:
			self.init_border() # Создаем отображение ячейки
		if rv.font_name:
			self.font_name = rv.font_name
		if rv.font_size:
			self.font_size = rv.font_size
		if rv.halign != '':
			self.halign = rv.halign
		return super().refresh_view_attrs(rv, index, data)

	def refresh_view_layout(self, rv, index, layout, viewport):
		# Обновляем rect
		x, y, = layout['pos']
		w,h = layout['size']
		self.rect.rectangle = (x, y, w, h)
		return super().refresh_view_layout(rv, index, layout, viewport)

	def on_minimum_height(self, instance, value):
		if not self.rv.auto_height:		# Если отключен режим автоподстройки высоты строк
			return
		if self.height == value:
			return
		# Если идет процесс обновления layout, то данные м.б. не верные, изменение высоты игнорируем
		if not self.rv.is_updating:
			self.rv.cell_min_height[self.index] = value

	def on_focus(self, instance, value):
		# Если идет процесс обновления layout, то фокус м.б. неверным, изменение фокуса игнорируем
		if not self.rv.is_updating:
			self.rv.data[self.index]['focus'] = value
		if value and not self._bubble:
			self._show_cut_copy_paste(self.pos, self, True)	#для нормального функционирования Bubble

	def on_text(self, instance, text):
		if self.parent and self.focus:
			self.rv.data[self.index]['text'] = text
	
	if kivy_version == '1.11.1':
		# Переопределенная функция TextInput. Для работы bubble при нажатии на handle
		def _handle_released(self, instance):
			sf, st = self._selection_from, self.selection_to
			if sf == st:
				return
			pos = (instance.right if instance is self._handle_left else instance.x, instance.top + self.line_height)
			pos = self.to_parent(*pos, relative=True)
			self._show_cut_copy_paste(pos, EventLoop.window)


"""
Класс отображения данных таблицы.
Для быстрого вывода большого количества данных используется RecycleView (он уже включает ScrollView)
"""
class DataTable(RecycleView):
	cell_min_height = DictProperty({})
	is_updating = BooleanProperty(False)
	auto_height = False		# Режим автоматической подстройки высоты строк
	font_name = ''
	font_size = ''
	halign = ''
	def __init__(self, *args, **kwargs):
		super().__init__(**kwargs)
		# Настраиваем прокрутку
		self.do_scroll = (True, True)   # Включить горизонтальную и вертикальную прокрутку
		self.effect_cls = ScrollEffect  # Отключить эффект прокрутки страницы за пределы границ контента
		self.bar_width = BAR_WIDTH              # Ширина полосы прокрутки
		self.scroll_type = ['bars', 'content']  # Тип прокрутки - по перемещиню контента и ползунка
		# Настраиваем Layout для размещения данных таблицы
		self.grid = RecycleGridLayout()
		self.grid.cols = 1 # default
		self.grid.padding = [0, 0, 2, 2] #Если с нумерацией строк. Без нее = [2, 0, 2, 2]
		self.grid.spacing = 0
		# Настраиваем размер layout чтобы работала прокрутка
		self.grid.size_hint = (None, None)
		self.grid.bind(minimum_size = self.grid.setter('size'))
		# Настраиваем размеры ячеек
		self.grid.default_size_hint = (1, 1)
		self.grid.col_default_width = COL_DEFAULT_WIDTH # Ширина столбцов по-умолчанию
		self.grid.row_default_height = ROW_DEFAULT_HEIGHT # Высота строк по-умолчанию
		self.add_widget(self.grid)
		# Настраиваем RecycleView для отображения данных
		self.viewclass = CellData   # Класс виджета для отображения данны
		if self.auto_height:
			# Создаем триггер обновления layout, чтобы все изменения обрабатывались за один раз
			self.init_trigger()

	def init_trigger(self, *largs):
		# Создаем триггер обновления layout, чтобы все изменения обрабатывались за один раз
		self.trigger_upd_grid = Clock.create_trigger(self._upd_grid)
		self.bind(data = self._set_rows_height_default)

	def on_cell_min_height(self, instance, value):
		""" При изменении высоты ячейки запускает процесс пересчета и обновления layout """
		self.trigger_upd_grid()

	def _upd_grid(self, *largs):
		""" Пересчитывает и обновляет layout """
		cols = self.grid.cols
		need_upd = False
		for row_num in self.grid.rows_minimum:
			row = [self.cell_min_height.get(row_num*cols+key) for key in range(cols)]
			max_height = max(row)
			if max_height != self.grid.rows_minimum[row_num]:
				self.grid.rows_minimum.update({row_num: max_height})
				need_upd = True
		if need_upd:
			self.refresh_from_layout()	# Обновляем layout
			self.is_updating = True		# Помечаем, что идет процесс обновления
			Clock.schedule_once(self._end_updating)

	def _end_updating(self, *largs):
		self.is_updating = False

	def _set_rows_height_default(self, instance, value):
		""" Устанавливает высоту строк по-умолчанию """
		length = len(value)
		if length > 0:
			row_count = int(length/self.grid.cols)
			self.cell_min_height = {i: 0 for i in range(length)}
			rows = {i: self.grid.row_default_height for i in range(row_count)}
			self.grid.rows_minimum = rows

""" 
Класс таблицы для отображения данных. Таблица состоит из двух частей: 
	ColHeaders - заголовки столбцов таблицы
	RowHeaders - заголовки строк таблицы
	Body - данные таблицы
"""
class WorkTable(BoxLayout):
	"""
	TODO: 
	1. В полях доп. опций д.б. изменение размера столбцов и строк. А т.ж. возможность цветовой заливки, копировать, вставить, скрыть строку/столбец
	2. Добавить возможность выбора страниц таблицы
	3. Сделать настройку цвета заголовочных полей и цвета поля данных таблицы. Оба этих параметра будут как цвета по-умолчанию.
	"""
	# Заголовки столбцов можно не задавать - тогда они автоматически будут нумероваться в буквенном обозначении.
	# Через функцию set_font можно задавать шрифт для всей таблицы (вызывать до заполнения таблицы).
	# Индивидуальный шрифт ячейки задавать в общем списке данных (data) 
	# Заголовки столбцов таблицы. Аналогично свойству data
	headers = ListProperty([])
	#Данные таблицы. Список из словарей (для каждой ячейки) в формате [{'text':данные}].
	#Можно задавать и другие свойства, например 'fill_color'.
	data = ListProperty([])
	#Индивидуальная ширина каждого столбца. Словарь в формате {номер столбца:ширина столбца}. Номера столбоцов начинаются с 0. Данные с ключом 'text' д.б. типа string.
	columns_width = DictProperty({})
	#Индивидуальная высота каждой строки. Словарь в формате {номер строки:высота строки}. Номера строк начинаются с 0.
	rows_height = DictProperty({})
	# Еденица измерения ширины строки (в символах или пикселях). Excell измеряет в количестве символов '0'
	col_width_unit = OptionProperty('characters+padding', options = ['characters+padding', 'characters', 'pixels'])
	row_height_unit = OptionProperty('rowcount*rowheight', options = ['rowcount*rowheight', 'rowcount', 'pixels'])
	cols = NumericProperty(0)
	auto_height = BooleanProperty(False)	# Режим автоматической подстройки высоты строк
	char_width_px = CHARCTER_WIDTH_PX	# Количество пикселей в одном симовле текста
	
	def __init__(self, **kwargs):
		event_mouse = EventMouseCursor()	# Событие изменения курсора мыши при наведении на виджет
		event_resize = EventResizeColRow()		# Событие изменения размера столбца
		event_resize.bind(on_resize_col = self.resize_col)
		event_resize.bind(on_resize_row = self.resize_row)
		super().__init__(**kwargs)
		self.orientation = 'vertical'
		# Добавляем названия столбцов таблицы
		self.ColHeaders = HeadTable(event_mouse, event_resize, location='horizontal')
		blh = BoxLayout(orientation='horizontal')
		blh.size_hint=(1, None)
		blh.height = ROW_DEFAULT_HEIGHT
		self.corner = Button(on_press=self.on_press_corner)
		self.corner.size_hint = (None, 1)
		self.corner.width = ROW_NAME_WIDTH
		self.corner.background_normal = ''
		blh.add_widget(self.corner)
		blh.add_widget(self.ColHeaders)
		self.add_widget(blh)
		# Добавляем нумерацию строк таблицы
		self.RowHeaders = HeadTable(event_mouse, event_resize, location='vertical')
		blb = BoxLayout(orientation='horizontal')
		blb.add_widget(self.RowHeaders)
		# Добавляем таблицу данных
		self.Body = DataTable()
		blb.add_widget(self.Body)
		self.add_widget(blb)
		# Делаем зависимую прокрутку по оси x для ColHeaders и Body
		self.ColHeaders.fbind('scroll_x', self.Body.setter('scroll_x'))
		self.Body.fbind('scroll_x', self.ColHeaders.setter('scroll_x'))
		# Делаем зависимую прокрутку по оси y для RowHeaders и Body
		self.RowHeaders.fbind('scroll_y', self.Body.setter('scroll_y'))
		self.Body.fbind('scroll_y', self.RowHeaders.setter('scroll_y'))
		
	def on_cols(self, instance, value):
		self.Body.grid.cols = value
	
	def on_headers(self, instance, value):
		cols = len(value)
		self.ColHeaders.data = value
		self.corner.background_color = DEFAULT_HEAD_BG_COLOR
		self.Body.grid.cols = cols

	def on_data(self, instance, value):
		# Чтобы данные отобразились в таблице, д.б. указано количество столбцов (определяется при загрузке headers)
		# Заполняются поля значениями по-умолчанию (если не указаны):
		for pos, data in enumerate(value):
			if 'background_color' not in data.keys():
				value[pos].update({'background_color': DEFAULT_BODY_BG_COLOR})
			if 'focus' not in data.keys():
				value[pos].update({'focus': False})
		# Если headers не загружен, заголовки будут с буквенным обозначением столбцов
		if not self.headers:
			self.headers = [{'text': self._num_char(c), 'background_color': DEFAULT_HEAD_BG_COLOR, 'focus': False} for c in range(self.cols)]
		# Задаем нумерацию строк
		lv = len(value)
		r_count = lv//self.cols+(lv%self.cols>0) #Количество строк
		self.RowHeaders.data = [{'text': str(r), 'background_color': DEFAULT_HEAD_BG_COLOR, 'focus': False} for r in range(1, r_count+1)]
		self.Body.data = value

	def on_columns_width(self, instance, value):
		# width = (character_count * character_width*density) + padding_left + padding_right # По правильному
		# Но Excel возвращает character_count вместе с pading -> использовать col_width_unit = 'characters+padding'
		# character_count - количество символов '0' в ширине
		# character_width = 7 (at symbol '0' in pixels) = kivy.core.text.Text(text='0').width
		# density - плотность экрана. М.б. на мобильных = 2; на ПК = 1. Учитывается в dp()
		# paddingleft + padding_right = отступы слева и справа от текста
		# value принимает словарь с шириной в пикселях для каждого столбца
		width = {}
		if self.col_width_unit == 'characters+padding':
			# Переводим ширину из количества символов '0'+padding в кол-во пикселей
			# Так выдает Excel
			for val in value.items():
				#width.update({val[0]: round(val[1] * dp(CHARCTER_WIDTH_PX))})
				width.update({val[0]: round(val[1] * dp(self.char_width_px))})
		elif self.col_width_unit == 'characters':
			# Переводим ширину из количества символов '0' в кол-во пикселей
			for val in value.items():
				#width.update({val[0]: round(val[1] * dp(CHARCTER_WIDTH_PX + PADDING[0] + PADDING[2]))})
				width.update({val[0]: round(val[1] * dp(self.char_width_px + PADDING[0] + PADDING[2]))})
		elif self.col_width_unit == 'pixels':
			width = value
		self.ColHeaders.grid.cols_minimum = width
		self.Body.grid.cols_minimum = width
		if self.data:
			self.ColHeaders.event_mouse.clear_listen_objects()	# Удаляем объекты события (при обновлении создаются новые)
			self.ColHeaders.refresh_from_layout()
			self.Body.refresh_from_layout()

	def on_rows_height(self, instance, value):
		height = {}
		if self.row_height_unit == 'rowcount*rowheight':
			# Переводим высоту из (количества строк)*(высоту строки) в кол-во пикселей
			# Так выдает Excel
			for val in value.items():
				height.update({val[0]: round(dp(val[1]*96/72))})
		elif self.row_height_unit == 'rowcount':
			# Переводим высоту из количества строк в кол-во пикселей
			for val in value.items():
				height.update({val[0]: round(pt(val[1] * CHARCTER_HEIGHT_PX))})
		elif self.row_height_unit == 'pixels':
			height = value
		self.Body.grid.rows_minimum = height
		self.RowHeaders.grid.rows_minimum = height
		if self.data:
			self.ColHeaders.event_mouse.clear_listen_objects()	# Удаляем объекты события (при обновлении создаются новые)
			self.ColHeaders.refresh_from_layout()
			self.Body.refresh_from_layout()

	def on_auto_height(self, instance, value):
		self.ColHeaders.auto_height = value
		self.Body.auto_height = value
		if value:
			self.ColHeaders.init_trigger()
			self.Body.init_trigger()

	def _calc_pix_in_char(self, font_name, font_size, char='0'):
		# Функция вычисляет размер текста в пикселях
		txt = TXT(text=char)
		txt.options['font_name'] = font_name
		txt.options['font_size'] = font_size
		txt.resolve_font_name()	# Разрешить имя шрифта
		txt.refresh()
		return txt.width, txt.height

	def set_font(self, font_name, font_size, path=None):
		# Функция устанавливает шрифт для всей рабочей части таблицы
		# Шрифт ищется в заранее определенном наборе каталогов
		# Если шрифт не найден - применяется стандартный
		# path - путь к пользовательскому ресурсу со шрифтами
		if path:
			# Добавляем path к ресурсам для поиска шрифтов
			resources.resource_add_path(path)
		font_file = font_name + '.ttf'
		if not resources.resource_find(font_file):	# возвращает путь к шрифту
			font_name = DEFAULT_FONT_NAME	# Шрифт по-умолчанию
		if type(font_size) == int or float:
			font_size = round(dp(font_size))
		self.Body.font_name = font_name
		self.Body.font_size = font_size
		self.char_width_px = self._calc_pix_in_char(font_name, 14)[0] #font_size)
		#print("Set FONT:", font_name, font_size, self.char_width_px)

	def set_alignment(self, halign):
		# Функция устанавливает выравнивание текста для всей таблицы
		# Пока доступно только выравнивание по горизонтали
		self.Body.halign = halign

	def resize_col(self, instance, touch):
		""" Функция начинает процедуру изменения размера виджета: захватывает нажатие и создает линию нового размера.
		Дальше по перемещению и отпусканию нажатия происходит перемещение линии и применение нового размера
		========= Можно добавить функцию скрытия ячеек. Если новый размер ячейки уходит в минус - скрыть ячейку. Только совместно с полноценной функцией скрыть/показать ячейку
		"""
		touch.grab(self)
		with self.canvas.after:
			Color(*RES_LINE_COLOR)
			touch.ud['line'] = Line(points=[touch.x, self.top-self.ColHeaders.grid.padding[1], touch.x, self.y], dash_offset = RES_DASH_OFFSET)
			touch.ud['type'] = 'vert'
			
	def resize_row(self, instance, touch):
		""" Функция начинает процедуру изменения размера виджета: захватывает нажатие и создает линию нового размера.
		Дальше по перемещению и отпусканию нажатия происходит перемещение линии и применение нового размера
		========= Можно добавить функцию скрытия ячеек. Если новый размер ячейки уходит в минус - скрыть ячейку. Только совместно с полноценной функцией скрыть/показать ячейку
		"""
		touch.grab(self)
		with self.canvas.after:
			Color(*RES_LINE_COLOR)
			touch.ud['line'] = Line(points=[self.x-self.RowHeaders.grid.padding[0], touch.y, self.right, touch.y], dash_offset = RES_DASH_OFFSET)
			touch.ud['type'] = 'hor'

	def on_touch_move(self, touch):
		""" Ослеживает захваченное перемещение и отрисовываем линию изменения размера"""
		if touch.grab_current is self:
			# Для вертикальной линии
			if touch.ud['type']=='vert':
				touch.ud['line'].points = [touch.x, self.top-self.ColHeaders.grid.padding[1], touch.x, self.y]
				# Если расширение выходит за пределы окна:
				if touch.x >= self.width-10:
					self.ColHeaders.scroll_x += self.ColHeaders.convert_distance_to_scroll(10, 0)[0]	# Увеличиваем поле просмотра
					if touch.ud['index'] in self.ColHeaders.view_adapter.views.keys():
						cell = self.ColHeaders.view_adapter.views[touch.ud['index']]	# Получаем экземпляр ячейки
						touch.ud['pos_x'] = cell.to_window(*cell.pos)[0]	# Записываем новую позицию x ячейки
			# Для горизонтальной линии
			if touch.ud['type']=='hor':
				touch.ud['line'].points = [self.x-self.RowHeaders.grid.padding[0], touch.y, self.right, touch.y]
				# Если расширение выходит за пределы окна:
				if touch.y <= self.y+10:
					self.RowHeaders.scroll_y -= self.RowHeaders.convert_distance_to_scroll(0, 10)[1]	# Увеличиваем поле просмотра
					if touch.ud['index'] in self.RowHeaders.view_adapter.views.keys():
						cell = self.RowHeaders.view_adapter.views[touch.ud['index']]	# Получаем экземпляр ячейки
						touch.ud['pos_y'] = cell.to_window(0, cell.top)[1]	# Записываем новую позицию y ячейки
	
	def on_touch_up(self, touch):
		""" По отпусканию захваченного нажатия расчитывает и применяет новый размер к виджету"""
		if touch.grab_current is self:
			# вычисляем и применяем заданный размер
			if touch.ud['type']=='vert':
				new_width = int(touch.x - touch.ud['pos_x'])
				#self.columns_width.update({touch.ud['index']: new_width})
				# Чтобы не тратить ресурсы на обработку еденицы измерения ширины столбца, записываем ширину в пикселях сразу в recycleview и обновляем
				self.ColHeaders.grid.cols_minimum[touch.ud['index']] = new_width
				self.Body.grid.cols_minimum[touch.ud['index']] = new_width
				self.ColHeaders.refresh_from_layout()
				self.ColHeaders.grid.goto_view(touch.ud['index'])
			elif touch.ud['type']=='hor':
				new_height = int(touch.ud['pos_y'] - touch.y)
				self.RowHeaders.grid.rows_minimum[touch.ud['index']] = new_height
				self.Body.grid.rows_minimum[touch.ud['index']] = new_height
				self.RowHeaders.grid.goto_view(touch.ud['index'])
				self.RowHeaders.refresh_from_layout()
				self.RowHeaders.grid.goto_view(touch.ud['index'])
			self.Body.refresh_from_layout()
			touch.ungrab(self)
			self.canvas.after.clear()	# Удаляем линию
		#else: super().on_touch_down(touch)
		
	def on_press_corner(self, instance):
		pass
		
	def _num_char(self, num):
		# Функция переводит числовое обозначение заголовка столбца в буквенное
		ch = ''
		while num >= 0:
			ch = str(chr(num%26+65))+ch
			num = num//26-1
		return ch



if __name__ == '__main__':
	from kivymd.app import MDApp
	from kivy.uix.boxlayout import BoxLayout
	from kivy.uix.button import Button
	class TestApp(MDApp):
		def build(self):
			bl = BoxLayout(orientation = 'vertical')
			bl.size = (800, 600)
			self.tb = WorkTable()
			b = Button(text= "Test Button", on_press=self.press_button)
			#b.background_color = (0.1, 0.1, 0.1, 0)
			self.ti = TextInput(text="text", size_hint=(1, 0.2), on_touch_down=self.press_text)
			#self.ti = Label(text="fgv")
			bl.add_widget(self.tb)
			bl.add_widget(self.ti)
			bl.add_widget(b)
			return bl
			
		def press_text(self, instance, touch):
			#print(f"TEST: {self.ti._bubble}")
			pass

		def press_button(self, instance):
			print("Test button pressed")
			print(f"BUTTON: {self.tb.headers}")
			print(f"BUTTON: {self.tb.data}")
			#print(dir(self.ti))
			head_color = (0.8,0.8,0.8,1)	# Цвет фона заголовка таблицы
			#table_color = (0.8,0.98,0.98,1)	# Цвет фона основной таблицы
			table_color = (1,1,1,1)
			# Настраиваем таблицу
			self.tb.set_font('Calibri', 11*96/72)	# Установили шрифт
			self.tb.col_width_unit = 'pixels'	# Установили еденицы измерения ширины столбцов
			self.tb.columns_width = {0: 350, 1: 159, 2:300, 3:350, 4:150, 6: 300}	#индивидуальная ширина столбцов
			self.tb.rows_height = {0: 50, 1: 59, 2:30, 3:35, 4:15, 6: 30}	#индивидуальная ширина столбцов
			#self.tb.auto_height = True
			# Добавляем наименования столбцов
			#self.tb.headers = [{'text': f"head {x}", 'background_color': head_color, 'focus': False} for x in range(6)]
			#self.tb.headers[1].update({'text': f"HEAD {1} \nkjdfklsjklfjs"}) #изменили наименование столбца
			#self.tb.headers[1]['text'] = f"HEAD {1}" #изменили свойство существующего поля
			#self.tb.headers.append({'text': f"HEAD {4}"})	# Добавили данные
			#self.tb.headers.append({'text': f"HEAD {5}"})
			#self.tb.data = [{'text': f"data {x}", 'focus': False} for x in range(18)]
			#self.tb.data[1].update({'text': f"DATA {1} \nkjdfklsjklfjs"}) #изменили свойство существующего поля
			#self.tb.data[11].update({'text': f"DATA {11} \nkjdfklsjklfjs"}) #изменили свойство существующего поля
			# Можно установить шрифт для каждой ячейки индивидуально (не рекомендуется):
			#data = [{'text': f"data {x}", 'background_color': table_color, 'focus': False,
			#		'font_name': 'Arial', 'font_size': '12.0pt'} for x in range(102)]
			data = [{'text': f"data {x}", 'background_color': table_color, 'focus': False} for x in range(1000)]
			data[0].update({'font_name': 'Roboto-Bold'})
			data[1].update({'font_name': 'Roboto-Bold'})
			data[5].update({'font_name': 'Roboto-Bold'})
			data[6].update({'font_name': 'Roboto-Bold'})
			#data = []
			#for x in range(1200):
			#	if (x+2)%6 != 0:
			#		data.append({'text': f"data {x}", 'background_color': table_color, 'focus': False})
				#else:
					#data.append({'text': None, 'background_color': table_color, 'focus': False})
			#data[1].update({'text': f"DATA {1} \nkjdfklsjklfjs fhjbvfgbbvf"}) #изменили свойство существующего поля
			#data[2].update({'text': None}) 
			#self.tb.headers[2].update({'text': f"Headers {2} \nkjdfklsjklfjs"}) #изменили свойство существующего поля
			data[11].update({'text': f"DATA {11} \nkjdfklsjklfjs"}) #изменили свойство существующего поля
			self.tb.cols = 7
			self.tb.data = data
			#for x in range(12):
			#	self.tb.data.append({'text': f"data2 {x}", 'background_color': table_color, 'focus': False, 'font_name': 'Arial', 'font_size': '11pt'})
			#print(f"BUTTON: {self.tb.headers}")
			#print(f"BUTTON: {dir(self.tb.ColHeaders.grid.children[0])}")
			
			print(f"BUTTON: END FUNC")



	TestApp().run()