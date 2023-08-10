from kivy.clock import Clock
from kivy.config import Config
from kivymd.uix.menu import MDDropdownMenu

Config.set("graphics", "resizable", 0)
Config.set("graphics", "width", 760)
Config.set("graphics", "height", 700)

from kivy.lang import Builder
from kivymd.app import MDApp
from kivymd.uix.filemanager import MDFileManager
from kivymd.toast import toast
import os
from parcer_xlsx import ParcerXlsxData
from threading import Thread


class FrdoStudent(MDApp):
    res = 0

    def __init__(self):
        super().__init__()
        self.manager_open = False

    def build(self):
        self.theme_cls.theme_style = "Dark"

        return Builder.load_file("frdostudent.kv")

    def ParcerLoad(self):
        file_1 = self.root.ids.file1
        file_2 = self.root.ids.file2
        spec = self.root.ids.spec
        file1 = list(file_1.split('\\'))
        file2 = list(file_2.split('\\'))
        if file_1 != "" and file_2 != "":
            if '.xlsx' in file1[-1] and '.xlsx' in file2[-1]:
                if spec != "":
                    self.root.transition.direction = 'left'
                    self.root.current = 'Load'
                    self.ParcerXlsxInfo(first_file=file_1.text, second_file=file_2.text, spec=spec.text)
                    file_1.text = ""
                    file_2.text = ""
                    spec.text = ""
                else:
                    toast('Не выбрана специальность')
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
            self.root.ids.file1.text = path

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
            self.root.ids.file2.text = path

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
            caller=self.root.ids.button,
            items=menu_items,
            width_mult=4,
        )

        menu.open()

    def text_spec(self, text):
        self.root.ids.spec.text = text

    def ParcerXlsxInfo(self, first_file, second_file, spec):
        th1 = Thread(target=ParcerXlsxData, args=(first_file, second_file, spec))
        th1.start()

        def callback(dt):
            if not th1.is_alive():
                self.res = 1
            if self.res == 1:
                self.res = 0
                event.cancel()
                self.data(spec=spec)
                self.root.transition.direction = 'left'
                self.root.current = 'Result'

        event = Clock.schedule_interval(callback, 2)

    def data(self, spec):
        self.root.ids.name.text = f"Файл сохранен под названием \n'Сведения студентов из ФРДО {spec}.xlsx'"
        self.root.ids.path.text = f"Файл сохранен по пути \n'C:\\Users\\{os.getlogin()}\\Documents\\Сведения студентов из ФРДО {spec}.xlsx'"


if __name__ == "__main__":
    FrdoStudent().run()