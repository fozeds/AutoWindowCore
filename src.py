from ctypes import windll
from os import system
from os.path import abspath, dirname
from time import sleep, time
import pygetwindow as gw
from keyboard import write
from pyautogui import doubleClick, locateCenterOnScreen, moveTo, typewrite
from pywinauto import application, findwindows

png = {
    1: "arvore1.png",
    2: "arvore2.png",
    ...: exemples,
    "r1n": "reverso1.png",
    ...: exemples,
    "3n": "arvore3n.png",
    "userpoint": "userpoint.png",
    "passwordpoint": "passwordpoint.png",
    "confirmpoint": "confirmpoint.png",
    "evadeerror": "evadeerror.png",
    "de": "de.png",
    "ate": "ate.png",
    "print": "imprimir.png",
    "continue": "continue.png",
    "close": "closewindow.png"
}


class TreeFinder():
    def __init__(self):
        self.folder_path = dirname(abspath(__file__).replace("src.py", ""))
        self.actual_window = None
        self.all_windows = gw.getAllTitles()
        self.programdir = r"\\dir\\file_or_lnk"
        self.login_image_dir = f"{self.folder_path}\\images\\{png["userpoint"]}"
        self.password_image_dir = f"{self.folder_path}\\images\\{png["passwordpoint"]}"
        self.confirm_image_dir = f"{self.folder_path}\\images\\{png["confirmpoint"]}"
        self.app = application.Application()    

    def search_and_open_window(
        self,
        window_title: str,
        has_login: bool = False,
        login_page: str = "",
        user: str = "",
        password: str = "",
        diff: tuple = ((0, 0), (0, 0), (0, 0)),
    ) -> None:
        """
        This method checks if the specified window is open.
        If not, it will search for the executable in the path defined
        in the program directory attribute.

        Args:
            window_title (str): The exact title of the window to be opened.
            has_login (bool): Indicates if login is required in the application
                            if it is not already open.
            login_page (str): The title of the login page, if there is a login.
            user (str): The username for authentication.
            password (str): The password for authentication.
            diff (tuple): A tuple ((x, y), (x, y), (x, y)) with three tuples
                        representing the differences needed for click between
                        the pixels at the center of the comparison images.
                        These can be used as optional parameters.
        """
        if window_title in self.all_windows:
            if gw.getActiveWindowTitle() != window_title:
                self.open_the_window(title=window_title)
        else:
            if self.programdir:
                # i really dont know why subprocess dont work with .lnk and serverless files.
                system(f"start {self.programdir}")
                sleep(2)
                if login_page:
                    timeout = time() + 20
                    while time() < timeout:
                        if gw.getActiveWindowTitle() == login_page:
                            break

                if has_login:
                    for difference, img, string in zip(
                        diff,
                        (
                            self.login_image_dir,
                            self.password_image_dir,
                            self.confirm_image_dir
                        ),
                        (user, password, None)
                    ):
                        x, y = locateCenterOnScreen(img)
                        moveTo(x + difference[0], y + difference[1])
                        doubleClick()
                        if string:
                            typewrite(string)
            else:
                self.alert_window("Window not found and directory not provided", "Error")

    def open_the_window(self, title: str) -> None:
        """
        Focuses the window with the specified title using the win32 API.

        Will not focus more than one window (in case of 2 or more with the same name).

        This method is specifically designed for Windows OS!

        Please refer to the documentation of other methods or classes used within this method.

        Args:
            title (str): Exact title of the window.
        """
        try:
            handles = findwindows.find_windows(title=title)
            if len(handles) > 1:
                raise ValueError('Multiple windows with this title were found')
            elif not handles:
                raise ValueError('Window with this title was not found')

            handle = handles[0]
        except ValueError as e:
            return self.alert_window(str(e), 'Error')

        self.app.connect(handle=handle)
        window = self.app.window(handle=handle)
        window.set_focus()
