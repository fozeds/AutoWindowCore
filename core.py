from ctypes import windll
from os import system
from os.path import abspath, dirname, join
from time import sleep, time

from pyautogui import doubleClick, locateCenterOnScreen, moveTo, typewrite
from pygetwindow import getActiveWindowTitle, getAllTitles
from pywinauto import application, findwindows

from files import png


class Core:
    """
    A class for automating window management and image-based interactions on a Windows operating system.

    Attributes:
        folder_path (str): The directory path where the current script is located.
        png (dict): A dictionary containing image file names associated with their keys.

    Methods:
        get_active_window_names():
            Returns a list of all currently active window titles.
        
        get_active_window_name():
            Returns the title of the currently active window.
        
        open_window(window_title: str):
            Focuses on the window with the specified title if it is active.
        
        open_software(programdir: str):
            Opens a software application located at the specified directory.
        
        wait_until_window_is_open(window: str, timeout: int = 20):
            Waits until the specified window becomes active or raises a TimeoutError.
        
        execute_image_based_write(
            auth_img_array: list|tuple,
            auth_str_array: list|tuple,
            diff: tuple = ((0, 0), (0, 0), (0, 0))
        ):
            Performs image-based interactions with the screen, such as clicking and typing, based on provided images and strings.
        
        search_open_and_auth(
            window_title: str,
            has_login: bool = False,
            login_page: str = "",
            user: str = "",
            password: str = "",
            diff: tuple = ((0, 0), (0, 0), (0, 0)),
            auth_image_path_array: list|tuple = ("", "", "")
        ) -> None:
            Checks if the specified window is open and attempts to open the software if it is not. If login is required, performs authentication.
        
        find_image_path(dict_key: str) -> str:
            Finds the path to an image file in the current working directory based on a key from the image dictionary.
        
        find_img_and_click(dict_key: str, difx: int = 0, dify: int = 0, delay: float = 0.2) -> None:
            Finds an image on the screen and performs a click at its location with optional offsets and delay.
        
        try_click_one_or_more(*dict_key_tuple: str) -> None:
            Attempts to find and click each specified image on the screen.
        
        wait_until_is_on_screen(key: str, timeout: int = 200, poll_interval: int = 2) -> None:
            Waits until a specified image appears on the screen or times out after the given duration.
        
        alert_window(message: str, title: str) -> None:
            Displays a Windows message box with the given message and title.
    """

    folder_path = dirname(abspath(__file__))
    png = png

    def __init__(self):
        self.actual_window = None
        self.app = application.Application()
    
    @staticmethod
    def get_active_window_names() -> list[str]:
        """
        Returns a list of all currently active window titles.

        Returns:
            list of str: A list containing the titles of all currently active windows.
        """
        return getAllTitles()

    @staticmethod
    def get_active_window_name() -> str:
        """
        Returns the title of the currently active window.

        Returns:
            str: The title of the currently active window.
        """
        return getActiveWindowTitle()

    def __open_the_window(self, title: str) -> None:
        """
        Focuses the window with the specified title using the win32 API.

        This method will focus on the first window found with the specified title. 
        If multiple windows have the same title, it will raise an error. If no 
        window with the given title is found, it will also raise an error.

        Note:
            This method is specifically designed for Windows OS.

        Args:
            title (str): The exact title of the window to focus on.

        Raises:
            ValueError: If multiple windows with the same title are found or if no window is found with the specified title.

        """
        try:
            # Find all window handles with the given title
            handles = findwindows.find_windows(title=title)
            
            if len(handles) > 1:
                raise ValueError('Multiple windows with this title were found')
            elif not handles:
                raise ValueError('Window with this title was not found')

            handle = handles[0]  # Use the first handle if there's only one

            # Connect to the window and set focus
            self.app.connect(handle=handle)
            window = self.app.window(handle=handle)
            window.set_focus()

        except ValueError as e:
            # Display an error message if there's an issue
            self.alert_window(str(e), 'Error')

    def open_window(self, window_title: str) -> None:
        """
        Focuses on the window with the specified title if it is currently active.

        This method checks if the given window title is among the active windows. 
        If it is, but the currently active window title does not match the given title, 
        it attempts to focus on the specified window. If the window title is not found 
        among the active windows, a ValueError is raised.

        Args:
            window_title (str): The title of the window to focus on.

        Raises:
            ValueError: If the window with the specified title is not found among the active windows.
        """
        # Get a list of all currently active window titles
        active_window_titles = self.get_active_window_names()
        
        if window_title in active_window_titles:
            # If the specified window is not currently active, focus on it
            if self.get_active_window_name() != window_title:
                self.__open_the_window(title=window_title)
        else:
            # Raise an error if the window title is not found among active windows
            raise ValueError("Cannot find window with the title '{}' among active windows.".format(window_title))
    
    def open_software(self, programdir: str):
        """
        Opens a software application located at the specified directory.

        This method attempts to open the software application specified by the given 
        directory path. If the `programdir` is empty or None, a ValueError is raised.

        Args:
            programdir (str): The directory path of the software application to open.

        Raises:
            ValueError: If the `programdir` is an empty string or None.
        """
        if programdir:
            system(f"start \"{programdir}\"")
        else:
            raise ValueError("Program directory is not set.")
    
    def wait_until_window_is_open(self, window: str, timeout: int = 20) -> None:
        """
        Waits until the specified window becomes active or raises a TimeoutError.

        This method repeatedly checks if the window with the specified title is the 
        currently active window. It will wait for up to the specified timeout period 
        before raising a TimeoutError if the window does not become active.

        Args:
            window (str): The title of the window to wait for.
            timeout (int, optional): The maximum amount of time (in seconds) to wait 
                                     for the window to become active. Defaults to 20 seconds.

        Raises:
            TimeoutError: If the specified window does not become active within the timeout period.
        """
        if not window:
            raise ValueError("Window title must be specified.")
        
        end_time = time.time() + timeout
        
        while time.time() < end_time:
            if self.get_active_window_name() == window:
                return
        
        raise TimeoutError(f"The window '{window}' did not become active within the timeout period of {timeout} seconds.")

    def execute_image_based_write(
        self, 
        auth_img_array: list[str], 
        auth_str_array: list[str], 
        diff: tuple[tuple[int, int], tuple[int, int], tuple[int, int]] = ((0, 0), (0, 0), (0, 0))
    ) -> None:
        """
        Performs image-based interactions with the screen, such as clicking and typing, based on provided images and strings.

        This method iterates over the provided images and strings, locating each image 
        on the screen and performing a click at the image's center. If a string is provided, 
        it types the string after clicking. The `diff` parameter specifies pixel offsets for 
        each image location to adjust where the click should occur.

        Args:
            auth_img_array (list[str]): A list of file paths to image files used for locating elements on the screen.
            auth_str_array (list[str]): A list of strings to be typed after each corresponding image is clicked.
            diff (tuple[tuple[int, int], tuple[int, int], tuple[int, int]], optional): Offsets (x, y) to adjust the click position for each image. Defaults to ((0, 0), (0, 0), (0, 0)).

        Raises:
            ValueError: If the lengths of `auth_img_array` and `auth_str_array` do not match.
        """
        if len(auth_img_array) != len(auth_str_array):
            raise ValueError("Length of image array must match length of string array.")
        
        for difference, img, string in zip(diff, auth_img_array, auth_str_array):
            sleep(0.5)  # Sleep for half a second to allow for any potential screen transitions
            x, y = locateCenterOnScreen(img, minSearchTime=2, confidence=0.9)
            moveTo(x + difference[0], y + difference[1])
            doubleClick()
            if string:
                typewrite(string)

    def search_open_and_auth(
        self,
        window_title: str,
        has_login: bool = False,
        login_page: str = "",
        user: str = "",
        password: str = "",
        diff: tuple[tuple[int, int], tuple[int, int], tuple[int, int]] = ((0, 0), (0, 0), (0, 0)),
        auth_image_path_array: list[str] | tuple[str, str, str] = ("", "", ""),
    ) -> None:
        """
        Checks if the specified window is open and attempts to open the software if it is not. If login is required,
        performs authentication using the provided images and credentials.

        This method first checks for the presence of the window with the specified title. If the window is not found,
        it attempts to open the software from a predefined directory. If a login is required and a login page title
        is provided, the method waits until the login page appears, then performs image-based authentication using 
        the provided credentials and images. The `diff` parameter allows for pixel offsets to adjust where interactions
        occur on the screen.

        Args:
            window_title (str): The exact title of the window to be checked.
            has_login (bool): Indicates whether login is required for the application if it is not already open. Defaults to False.
            login_page (str): The title of the login page, if there is a login screen. Defaults to an empty string.
            user (str): The username for authentication, if login is required. Defaults to an empty string.
            password (str): The password for authentication, if login is required. Defaults to an empty string.
            diff (tuple[tuple[int, int], tuple[int, int], tuple[int, int]], optional): A tuple ((x, y), (x, y), (x, y)) with three tuples representing pixel offsets needed for clicks between the center points of comparison images. Defaults to ((0, 0), (0, 0), (0, 0)).
            auth_image_path_array (list[str] | tuple[str, str, str], optional): A list or tuple containing file paths to images used for authentication. Defaults to ("", "", "").

        Raises:
            ValueError: If the window cannot be found and the software cannot be opened, or if there is an error
                        while attempting to open the software or perform the login.
        """
        auth_str_array = (user, password, None)
        try:
            self.open_window(window_title)
        except ValueError as open_window_error:
            try:
                self.open_software(self.programdir)
            except Exception as e:
                raise ValueError(f"Cannot find window: '{open_window_error}' and cannot open software: '{e}'")
            if login_page:
                self.wait_until_window_is_open(login_page)
                if has_login:
                    self.execute_image_based_write(auth_image_path_array, auth_str_array, diff)
            else:
                self.alert_window("Window not found and directory not provided", "Error")

    @classmethod
    def find_image_path(cls, dict_key: str) -> str:
        """
        Finds the path to an image file in the current working directory based on a key from the image dictionary.

        This class method searches for the image file path associated with the provided key in the `png` dictionary.
        The image path is constructed using the `folder_path` and `"images"` subdirectory, combined with the 
        filename retrieved from the `png` dictionary. If the provided key does not exist in the dictionary, 
        a `ValueError` is raised.

        Args:
            dict_key (str): The key used to look up the image file path in the `png` dictionary.

        Returns:
            str: The full path to the image file.

        Raises:
            ValueError: If the `dict_key` is not found in the `png` dictionary.

        Example:
            >>> MyClass.find_image_path('logo')
            '/path/to/current/directory/images/logo.png'
        """
        if dict_key not in cls.png:
            raise ValueError(f"Key '{dict_key}' not found in image dictionary")
        return join(cls.folder_path, "images", cls.png[dict_key])

    def find_img_and_click(
        self,
        dict_key: str,
        difx: int = 0,
        dify: int = 0,
        delay: float = 0.2
    ) -> None:
        """
        Finds an image on the screen and performs a click at its location with optional offsets and delay.

        This method uses the provided `dict_key` to locate the path of an image file through the `find_image_path` method.
        It then finds the center of the image on the screen and moves the cursor to that location, applying optional 
        pixel offsets (`difx` and `dify`). After moving to the location, it waits for a specified delay before performing 
        a double-click. Another delay is applied after the click to ensure that the action is properly registered.

        Args:
            dict_key (str): The key used to look up the image file path in the image dictionary.
            difx (int, optional): The horizontal offset (in pixels) to apply to the click position. Defaults to 0.
            dify (int, optional): The vertical offset (in pixels) to apply to the click position. Defaults to 0.
            delay (float, optional): The delay (in seconds) to wait before and after clicking. Defaults to 0.2 seconds.

        Raises:
            FileNotFoundError: If the image file specified by `dict_key` cannot be found.
            ValueError: If the image cannot be located on the screen.

        Example:
            >>> instance.find_img_and_click('button', difx=10, dify=-5, delay=0.5)
            # Finds the image associated with 'button', moves to (10, -5) offset from the center, and clicks with a 0.5-second delay.
        """
        place = Core.find_image_path(dict_key=dict_key)
        x, y = locateCenterOnScreen(place, minSearchTime=1, confidence=0.9)
        moveTo(x + difx, y + dify)
        sleep(delay)
        doubleClick()
        sleep(delay)

    def try_click_one_or_more(self, *dict_key_tuple: str) -> None:
        """
        Attempts to find and click each specified image on the screen.

        This method iterates over the provided image keys, attempting to locate and click each image using the 
        `find_img_and_click` method. If an image cannot be found or if an error occurs during the click attempt, 
        the method catches the exception and continues with the next image key.

        Args:
            *dict_key_tuple (str): One or more keys used to look up image file paths in the image dictionary. Each key corresponds to an image to be found and clicked on the screen.

        Example:
            >>> instance.try_click_one_or_more('button1', 'button2', 'button3')
            # Attempts to find and click images associated with 'button1', 'button2', and 'button3' in sequence.
        """
        for key in dict_key_tuple:
            try:
                self.find_img_and_click(key)
            except Exception:
                pass

    def wait_until_is_on_screen(
        self,
        key: str,
        timeout: int = 200,
        poll_interval: int = 2
    ) -> None:
        """
        Waits until a specified image appears on the screen or times out after the given duration.

        This method repeatedly attempts to find and click the image associated with the provided key. It waits 
        for a specified duration (`timeout`) and checks for the image's presence at regular intervals (`poll_interval`).
        If the image is found within the timeout period, the method performs a click and returns. If the timeout is reached 
        without finding the image, the method stops attempting and returns `None`.

        Args:
            key (str): The key used to look up the image file path in the image dictionary. This image is what the method 
                    is waiting to appear on the screen.
            timeout (int, optional): The maximum amount of time (in seconds) to wait for the image to appear. Defaults to 200 seconds.
            poll_interval (int, optional): The interval (in seconds) between each check for the image. Defaults to 2 seconds.

        Returns:
            None: The method does not return a value but will attempt to click the image if it appears before the timeout.

        Example:
            >>> instance.wait_until_is_on_screen('loading_spinner', timeout=100, poll_interval=5)
            # Waits up to 100 seconds, checking every 5 seconds, for the image associated with 'loading_spinner' to appear on the screen.
        """
        end_time = time() + timeout
        while time() < end_time:
            try:
                self.find_img_and_click(key)
                return
            except:
                sleep(poll_interval)

    @staticmethod
    def alert_window(message: str, title: str) -> None:
        """
        Displays a Windows message box with the given message and title.

        This static method uses the Windows API to create and display a message box with a specified message and title.
        It is intended for displaying alerts or notifications to the user within a Windows environment.

        Args:
            message (str): The text to be displayed in the message box.
            title (str): The title of the message box window.

        Example:
            >>> MyClass.alert_window("Operation completed successfully.", "Success")
            # Displays a message box with the message "Operation completed successfully." and the title "Success".
        """
        windll.user32.MessageBoxW(0, message, title, 0)