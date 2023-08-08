from kivy.config import Config
from kivymd.uix.menu import MDDropdownMenu

Config.set("graphics", "resizable", 0)
Config.set("graphics", "width", 760)
Config.set("graphics", "height", 700)

from kivy.clock import Clock
from kivy.lang import Builder
from kivy.uix.screenmanager import ScreenManager
from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.filemanager import MDFileManager
from kivymd.toast import toast
from kivy.core.window import Window
import os
from parcer_xlsx import ParcerXlsxData


Builder.load_file("frdostudent.kv")


class MainWindow(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        Window.bind(on_keyboard=self.events)
        self.manager_open = False

    def ParcerLoad(self):
        loadwin = LoadWindow()
        file_1 = self.ids.file1.text
        file_2 = self.ids.file2.text
        spec = self.ids.spec.text
        file1 = list(file_1.split('\\'))
        file2 = list(file_2.split('\\'))
        if file_1 != "" and file_2 != "":
            if '.xlsx' in file1[-1] and '.xlsx' in file2[-1]:
                self.manager.transition.direction = 'left'
                self.manager.current = 'Load'
                loadwin.ParcerXlsxInfo(first_file=file_1, second_file=file_2, spec=spec)
            else:
                toast("Файл(ы) не формата Excel")
        else:
            toast("Не выбраны файл(ы)")

    def open_file_1(self):
        def exit_manager(*args):
            self.manager_open = False
            file_manager.close()

        def select_path(path: str):
            exit_manager()
            self.ids.file1.text = path

        file_manager = MDFileManager(
            exit_manager=exit_manager, select_path=select_path
        )

        file_manager.show(os.path.expanduser("~"))

    def open_file_2(self):
        def exit_manager(*args):
            self.manager_open = False
            file_manager.close()

        def select_path(path: str):
            exit_manager()
            self.ids.file2.text = path

        file_manager = MDFileManager(
            exit_manager=exit_manager, select_path=select_path
        )

        file_manager.show(os.path.expanduser("~"))

    def spec_select(self):

        cabinet = [
            "Право и организация социального обеспечения",
            "Поварское и кондитерское дело",
            "Экономика и бухгалтерский учет (по отраслям)",
            "Информационные системы и программирование",
            "Защита в чрезвычайных ситуациях",
            "Ветеринария",
            "Павоохранительная деятельность",
            "Управление качеством продукции и услуг (по отраслям)",
                   ]

        menu_items = [
            {
                "text": f"{i}",
                "viewclass": "OneLineListItem",
                "on_release": lambda x=f"{i}": self.text_spec(x),
            } for i in cabinet
        ]

        menu = MDDropdownMenu(
            caller=self.ids.button,
            items=menu_items,
            width_mult=4,
        )

        menu.open()

    def text_spec(self, text):
        self.ids.spec.text = text


class LoadWindow(MDScreen):
    def ParcerXlsxInfo(self, first_file, second_file, spec):
        ParcerXlsxData(file1=first_file, file2=second_file, spec=spec)


class ResultFrdoWindow(MDScreen):
    pass


class FrdoStudent(MDApp):
    def build(self):
        self.theme_cls.theme_style = "Dark"
        sm = ScreenManager()
        sm.add_widget(MainWindow())
        sm.add_widget(LoadWindow())
        sm.add_widget(ResultFrdoWindow())

        return sm


if __name__ == "__main__":
    FrdoStudent().run()