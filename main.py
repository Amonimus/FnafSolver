import logging
import os
import threading
import time
import warnings
from typing import List
import sys
import keyboard
import pyautogui
import pywinauto
import win32gui
from pywinauto import mouse
from pywinauto.application import Application


class KeyboardEvent(threading.Thread):
    def run(self):
        while not keyboard.is_pressed('q') and not keyboard.is_pressed('Esc'):
            pass
        logging.error("Abort!")
        os._exit(1)


class FnafSolver:

    def __init__(self):
        self.game_path: str = "W:/Program Files (x86)/Steam/steamapps/common/Five Nights at Freddy's/FiveNightsatFreddys.exe"
        self.game_name: str = "Five Nights at Freddy's"
        self.fullscreen: bool = False
        self.app: Application = None
        self.window: pywinauto.application.WindowSpecification = None
        self.variables: dict = {
            "window_position": (0, 0),
            "side": "left",
            "left_light": False,
            "left_door": False,
            "right_light": False,
            "right_door": False,
            "see_chica": False,
            "see_bonnie": False,
            "foxy_stage": 0,
        }
        self.screen_size = (1280, 720)
        self.click_positionts: dict = {
            # Game res in widnowed is 1280x720
            "center": (64 * 9, 64 * 6),
            "center_bottom": (64 * 9, 64 * 9),
            "left": (8, 64 * 6),
            "right": (64 * 20, 64 * 6),
            "left_door": (64, 64 * 5),
            "left_light": (64, 64 * 7),
            "right_door": (64 * 19, int(64 * 5.5)),
            "right_light": (64 * 19, 64 * 7),
            "camera": (64 * 9, 64 * 11),
            "bonnie_check": (498, 300),
            "chica_check": (860, 268),
        }
        self.regions: dict = {
            "left_half": (0, 0, 64 * 6, 64 * 11),
            "right_half": (64 * 6, 0, 64 * 6, 64 * 11),
            "left_door": (4, 280, 100, 150),
            "right_door": (1180, 280, 100, 150),
            "left_light": (4, 390, 100, 150),
            "right_light": (1180, 390, 100, 150),
            "bonnie_lookup": (70, 200, 450, 350),
            "bonnie_lookup_small": (400, 200, 200, 300),
            "chica_lookup": (780, 180, 120, 250),
        }
        self.pixel_reference: dict = {
            "black": (0, 0, 0),
            "white": (255, 255, 255),
            "red": (173, 0, 0),
            "gray": (103, 103, 103),
            "bright": (184, 189, 208),
        }
        self.image_dict: dict = {
            "custom_night": "custom_night.png",
            "right_light_active": "right_light_active.png",
            "right_light_inactive": "right_light_inactive.png",
            "left_light_active": "left_light_active.png",
            "left_light_inactive": "left_light_inactive.png",
            "right_door_active": "right_door_active.png",
            "right_door_inactive": "right_door_inactive.png",
            "left_door_active": "left_door_active.png",
            "left_door_inactive": "left_door_inactive.png",
            "foxy_stage_1": "foxy_stage_1.png",
            "right_hall_chica": "right_hall_chica.png",
            "left_hall_bonnie": "left_hall_bonnie.png",
            "left_door_bonnie": "left_door_bonnie.png",
        }
        self.camera_spots: dict = {
            "1A": "B",
            "1B": " ",
            "1C": " ",
            "2A": " ",
            "2B": " ",
            "3": " ",
            "4A": " ",
            "4B": " ",
            "5": " ",
            "6": " ",
            "7": " ",
            "O": " ",
        }
        # Timings.after_clickinput_wait = 0.0001
        # Timings.after_setcursorpos_wait = 0.0001
        # win32gui.GetDoubleClickTime = self.click_time_callable

    def click_time_callable(self):
        return 0.0001

    def launch_game(self):
        logging.warning("Launching the game.")
        self.app = Application().start(self.game_path)
        time.sleep(6)

    def check_game_open(self) -> bool:
        open_wins: List[str] = [w.title for w in pyautogui.getAllWindows()]
        # logging.debug(f"Open apps: {open_wins}")
        return self.game_name in open_wins

    def get_active_window_name(self) -> str:
        active_window: str = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        logging.debug(f"Current window: {active_window}.")
        return active_window

    def get_app_from_running(self):
        logging.debug("Setting up app handle.")
        self.app = Application().connect(title=self.game_name)
        handles: List[int] = pywinauto.findwindows.find_windows(title=self.game_name)
        # logging.debug(f"Related window handles: {handles}")
        for handle in handles:
            # logging.debug(f"Connect: {handle}")
            self.app.connect(handle=handle)

    def get_main_window(self) -> pywinauto.application.WindowSpecification:
        # logging.debug(f"App windows: {self.app.windows()}")
        if self.window is None:
            self.window = pywinauto.application.WindowSpecification = self.app.window(title=self.game_name, visible_only=False)
        # logging.debug(f"App top window: {self.window.criteria}")
        return self.window

    def maximize_window(self):
        control = self.get_main_window()
        logging.debug(f"Unwrapping window.")
        control.minimize()
        control.restore()
        logging.debug(f"Game active? {control.is_active()}.")
        control.wait('visible')
        logging.debug(f"Game visible? {control.is_visible()}.")

    def open_game(self):
        if not self.check_game_open():
            logging.warning("Game is not running.")
            self.launch_game()
            need_restart = False
        else:
            logging.warning("Game is running.")
            need_restart = True

        if self.app is None:
            self.get_app_from_running()

        if self.get_active_window_name() != self.game_name:
            # Doesn't work if the game was clicked out instead of Alt+Tab
            self.maximize_window()
            if self.get_active_window_name() != self.game_name:
                raise Exception("Can't open game")

        self.get_window_position()
        logging.warning("Game OK.")
        time.sleep(2)
        # if need_restart:
        #     self.return_to_title()

    def seek_image(self, image: str, region: tuple = None) -> bool:
        if region:
            region = (
                self.variables["window_position"][0] + region[0],
                self.variables["window_position"][1] + region[1],
                self.variables["window_position"][0] + region[0] + region[2],
                self.variables["window_position"][1] + region[1] + region[3],
            )
            # logging.debug(f"Checking region: {region}.")
        try:
            location = pyautogui.locateOnScreen(image, region=region, confidence=0.9, grayscale=False)
            # logging.debug(f"See image {image} at: {location}.")
            return True
        except pyautogui.ImageNotFoundException:
            # logging.warning(f"Image {image} not found!")
            return False

    def mouse_action(self, coords: tuple, click: bool = False):
        # logging.warning(f"Intended coors: {coords}")
        coords = (self.variables["window_position"][0] + coords[0], self.variables["window_position"][1] + coords[1])
        if not self.fullscreen:
            coords = (coords[0], coords[1] + 32)
        # logging.warning(f"Offset coors: {coords}")
        if click:
            mouse.click(coords=coords)
        else:
            mouse.move(coords=coords)

    def get_pixel(self, orig_coords: tuple) -> tuple:
        coords = (self.variables["window_position"][0] + orig_coords[0], self.variables["window_position"][1] + orig_coords[1])
        if not self.fullscreen:
            coords = (coords[0], coords[1] + 32)
        pixel = pyautogui.pixel(coords[0], coords[1])
        # mouse.move((coords[0], coords[1]))
        # logging.debug(f"Pixel at position {orig_coords} is {pixel}")
        return pixel

    def validate_color(self, orig_coords: tuple, color: tuple, tolerance: int = 20) -> bool:
        coords = (self.variables["window_position"][0] + orig_coords[0], self.variables["window_position"][1] + orig_coords[1])
        if not self.fullscreen:
            coords = (coords[0], coords[1] + 32)
        # pixel = pyautogui.pixel(coords[0], coords[1])
        check = pyautogui.pixelMatchesColor(coords[0], coords[1], color, tolerance=tolerance)
        # mouse.move((coords[0], coords[1]))
        # logging.debug(f"Pixel {pixel} at position {orig_coords} is {color}? {check}")
        return check

    def toggle_monitor(self):
        self.mouse_action(self.click_positionts["camera"])
        self.mouse_action(self.click_positionts["center_bottom"])

    def confirm_screen(self):
        logging.debug("Checking screen.")
        pos = "unknown"
        if self.seek_image(self.image_dict["right_door_inactive"], self.regions["right_door"]):
            pos = "right"
        elif self.seek_image(self.image_dict["left_door_inactive"], self.regions["left_door"]):
            pos = "left"
        elif self.seek_image(self.image_dict["right_light_inactive"], self.regions["right_light"]):
            pos = "right"
        elif self.seek_image(self.image_dict["left_light_inactive"], self.regions["left_light"]):
            pos = "left"
        elif self.seek_image(self.image_dict["right_door_active"], self.regions["right_door"]):
            pos = "right"
        elif self.seek_image(self.image_dict["left_door_active"], self.regions["left_door"]):
            pos = "left"
        elif self.seek_image(self.image_dict["right_light_active"], self.regions["right_light"]):
            pos = "right"
        elif self.seek_image(self.image_dict["left_light_active"], self.regions["left_light"]):
            pos = "left"
        self.variables["side"] = pos
        logging.info(f"Current position: {pos}")
        return pos

    def launch_night(self):
        # Assumes Night 7 is unlocked and is saved as 4/20
        logging.info("Starting...")
        self.mouse_action((300, 640), True)
        time.sleep(1)
        # self.set_custom_night()
        self.mouse_action((1100, 640), True)
        attempt = 0
        logging.info("Awaiting the Office.")
        self.mouse_action(self.click_positionts["center"])
        while (not self.seek_image(self.image_dict["left_light_inactive"], self.regions["left_light"])
               and not self.seek_image(self.image_dict["right_light_inactive"], self.regions["right_light"])):
            attempt += 1
            logging.debug(f"Attempt {attempt}")
            if attempt >= 20:
                logging.error("Timeout!")
                return False
            logging.info("Night hasn't stated?")
            time.sleep(1)
        logging.info("Night stared.")
        return True

    def set_custom_night(self):
        if not self.seek_image(self.image_dict["custom_night"], (80, 400, 1120, 150)):
            logging.info("Custom night is not set. Wait...")
            mass_click_delay = 0.5
            for i in range(20):
                self.mouse_action((300, 500), True)
                time.sleep(mass_click_delay)
            for i in range(20):
                self.mouse_action((580, 500), True)
                time.sleep(mass_click_delay)
            for i in range(20):
                self.mouse_action((860, 500), True)
                time.sleep(mass_click_delay)
            for i in range(20):
                self.mouse_action((1150, 500), True)
                time.sleep(mass_click_delay)
        logging.info("Custom night is set.")

    def get_window_position(self):
        rect: pywinauto.win32structures.RECT = bot.get_main_window().rectangle()
        logging.debug(f"Window position: {(rect.left, rect.top)}")
        self.variables["window_position"] = (rect.left, rect.top)
        return self.variables["window_position"]

    def return_to_title(self):
        logging.warning("Restart.")
        keyboard.press('F2')
        time.sleep(5)

    def print_room_ascii(self):
        map = f"""
   [{self.camera_spots['1A']}][ ][ ]
[{self.camera_spots['5']}][{self.camera_spots['1B']}]   [ ][{self.camera_spots['7']}]
[{self.camera_spots['1C']}][ ]   [ ]
[{self.camera_spots['3']}][{self.camera_spots['2A']}]   [{self.camera_spots['4A']}][{self.camera_spots['6']}]
   [{self.camera_spots['2B']}][ ][{self.camera_spots['4B']}]
        """
        logging.info(map)

    def action(self, command: str):
        logging.debug(f"Variables: {self.variables}")
        match command:
            case "validate_light_left":
                logging.info("Validating left lights.")
                if not bot.validate_color(bot.click_positionts["left_light"], bot.pixel_reference["bright"], tolerance=30):
                    logging.info("Left is not active.")
                    self.variables["right_left"] = False
                else:
                    logging.info("Left is active.")
                    self.variables["right_left"] = True
            case "validate_light_right":
                logging.info("Validating right lights.")
                if not bot.validate_color(bot.click_positionts["left_right"], bot.pixel_reference["bright"], tolerance=30):
                    logging.info("Right is not active.")
                    self.variables["right_light"] = False
                else:
                    logging.info("Right is active.")
                    self.variables["right_light"] = True
            case "validate_door_left":
                logging.info("Validating left door.")
                if bot.validate_color(bot.click_positionts["left_door"], bot.pixel_reference["red"]):
                    logging.info("Left is open.")
                    self.variables["left_door"] = False
                else:
                    logging.info("Left is closed.")
                    self.variables["left_door"] = True
            case "validate_door_right":
                logging.info("Validating right door.")
                if bot.validate_color(bot.click_positionts["right_door"], bot.pixel_reference["red"]):
                    logging.info("Right is open.")
                    self.variables["right_door"] = False
                else:
                    logging.info("Right is closed.")
                    self.variables["right_door"] = True
            case "pan_left":
                logging.info("Panning left...")
                if self.confirm_screen() == "right":
                    logging.info("Panning left.")
                    self.mouse_action(self.click_positionts["left"])
                    time.sleep(0.3)
            case "pan_right":
                logging.info("Panning right...")
                if self.confirm_screen() == "left":
                    logging.info("Panning right.")
                    self.mouse_action(self.click_positionts["right"])
                    time.sleep(0.3)
            case "lock_left":
                logging.info("Locking left...")
                self.action("pan_left")
                self.action("validate_door_left")
                if not self.variables["left_door"]:
                    logging.info("Clicking left door.")
                    self.mouse_action(self.click_positionts["left_door"], True)
            case "lock_right":
                logging.info("Locking right...")
                self.action("pan_right")
                self.action("validate_door_right")
                if not self.variables["right_door"]:
                    logging.info("Clicking right door.")
                    self.mouse_action(self.click_positionts["right_door"], True)
            case "unlock_left":
                logging.info("Unlocking left.")
                self.action("pan_left")
                self.action("validate_door_left")
                if self.variables["left_door"]:
                    logging.info("Clicking left door.")
                    self.mouse_action(self.click_positionts["left_door"], True)
            case "unlock_right":
                logging.info("Unlocking right.")
                self.action("pan_right")
                self.action("validate_door_right")
                if self.variables["right_door"]:
                    logging.info("Clicking right door.")
                    self.mouse_action(self.click_positionts["right_door"], True)
            case "light_left":
                logging.info("Testing left light...")
                self.action("pan_left")
                self.mouse_action(self.click_positionts["left_light"], True)
                # self.action("validate_light_left")
                # if not self.variables["left_light"]:
                #     self.mouse_action(self.click_positionts["left_light"], True)
                #     time.sleep(0.1)
                #     self.mouse_action(self.click_positionts["left_light"], True)
            case "light_right":
                logging.info("Testing right light...")
                self.action("pan_right")
                self.mouse_action(self.click_positionts["right_light"], True)
            case "toogle_monitor":
                logging.info("Toggle monitor.")
                self.toggle_monitor()
            case "see_chica":
                logging.info("Validating Chica.")
                self.get_pixel((845, 220))
                if not self.validate_color(bot.click_positionts["chica_check"], self.pixel_reference["black"], tolerance=10):
                    logging.warning("Can see Chica.")
                    self.variables["see_chica"] = True
                else:
                    logging.info("Can't see Chica.")
                    self.variables["see_chica"] = False
            case "see_bonnie":
                logging.info("Validating Bonnie.")
                self.get_pixel((442, 230))
                # if (self.seek_image(self.image_dict["left_hall_bonnie"], self.regions["bonnie_lookup_small"])
                #         or self.seek_image(self.image_dict["left_door_bonnie"], self.regions["bonnie_lookup"])):
                if self.validate_color(self.click_positionts["bonnie_check"], self.pixel_reference["black"], tolerance=20):
                    logging.warning("Can see Bonnie.")
                    self.variables["see_bonnie"] = True
                else:
                    logging.info("Can't see Bonnie.")
                    self.variables["see_bonnie"] = False
            case "check_chica":
                logging.info("Looking for Chica...")
                self.action("light_right")
                self.action("see_chica")
                if self.variables["see_chica"]:
                    self.action("lock_right")
                else:
                    self.action("unlock_right")
                self.action("light_right")
            case "check_bonnie":
                logging.info("Looking for Bonnie...")
                self.action("light_left")
                self.action("see_bonnie")
                if self.variables["see_bonnie"]:
                    self.action("lock_left")
                else:
                    self.action("unlock_left")
                self.action("light_left")
            case "check_foxy":
                logging.info("Looking for Foxy.")
            case _:
                logging.error(f"Unknown action {command}")
                raise Exception(f"Unknown action {command}")

    def start_loop(self):
        try:
            start_time = time.time()
            night_duration = 435  # 7 min 15 s
            if not self.launch_night():
                Exception("Night Failed to start!")
            while time.time() < start_time + night_duration:
                self.action("check_chica")
                self.action("check_bonnie")
                # self.action("check_foxy")
                # time.sleep(2)
            logging.warning("Assuming night is over.")
        except Exception as e:
            logging.error(f"Exception in main loop: {e}")
            os._exit(1)


class LoggingFormatter(logging.Formatter):
    fmt = "%(asctime)s - %(levelname)s - %(message)s"
    FORMATS = {
        logging.DEBUG: "\x1b[36;20m" + fmt + "\x1b[0m",
        logging.INFO: "\x1b[38;20m" + fmt + "\x1b[0m",
        logging.WARNING: "\x1b[33;20m" + fmt + "\x1b[0m",
        logging.ERROR: "\x1b[31;20m" + fmt + "\x1b[0m",
        logging.CRITICAL: "\x1b[31;20m" + fmt + "\x1b[0m"
    }

    def format(self, record):
        log_fmt = self.FORMATS.get(record.levelno)
        formatter = logging.Formatter(log_fmt)
        return formatter.format(record)


def setup_logging():
    warnings.filterwarnings("ignore")
    log_handler = logging.StreamHandler(sys.stdout)
    log_handler.setLevel(logging.DEBUG)
    log_handler.setFormatter(LoggingFormatter())
    logging.basicConfig(handlers=[log_handler], level=logging.DEBUG)


if __name__ == '__main__':
    setup_logging()
    key_watch: KeyboardEvent = KeyboardEvent()
    try:
        bot: FnafSolver = FnafSolver()
        key_watch.start()
        bot.open_game()
        bot.start_loop()
        key_watch.join()
    except Exception:
        pass
        # key_watch.join()