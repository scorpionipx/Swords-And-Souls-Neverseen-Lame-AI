import cv2

import keyboard

import numpy as np

from PIL import ImageGrab

from time import sleep

from threading import Thread

from win32gui import FindWindow, GetWindowRect

from win32com import client


GAME_WINDOW_TITLE = 'Swords & Souls Neverseen'


class SwordsAndSoulsNeverseenAI:
    """SwordsAndSoulsNeverseenAI

    """
    def __init__(self):
        """__init__

            Initialization.
        """
        self.game_window = None
        self.__get_game_window__()

        self.keyboard_client = client.Dispatch("WScript.Shell")

        self.left_allowed = True
        self.top_allowed = True
        self.right_allowed = True

        self.melee_training_running = False
        self.__enable_key_shortcut__()

    def __melee_training__(self):
        """melee_training

        :return:
        """
        ss = ImageGrab.grab(bbox=self.area_of_interest)
        # ss.show()
        ss = cv2.cvtColor(np.array(ss), cv2.COLOR_BGR2GRAY)

        x = 162
        y = 4
        w = 50
        h = 60

        top_hit_box = ss[y:y+h, x:x+w]

        # cv2.imshow('top grey', top_hit_box)
        # cv2.waitKey()

        ret, top_hit_box = cv2.threshold(top_hit_box, 120, 255, cv2.THRESH_BINARY)

        # cv2.imshow('top binary', top_hit_box)
        # cv2.waitKey()

        white_pixels = cv2.countNonZero(top_hit_box)
        # print(f'top wp: {white_pixels}')

        if white_pixels < 2600 and self.top_allowed:
            self.disable_top()
            self.keyboard_client.SendKeys('w')

        x = 10
        y = 162
        w = 60
        h = 50

        left_hit_box = ss[y:y+h, x:x+w]

        # cv2.imshow('left grey', left_hit_box)
        # cv2.waitKey()

        ret, left_hit_box = cv2.threshold(left_hit_box, 120, 255, cv2.THRESH_BINARY)

        # cv2.imshow('left binary', left_hit_box)
        # cv2.waitKey()

        white_pixels = cv2.countNonZero(left_hit_box)
        # print(f'left wp: {white_pixels}')

        if white_pixels < 2600 and self.left_allowed:
            self.disable_left()
            self.keyboard_client.SendKeys('a')

        x = 310
        y = 162
        w = 60
        h = 50

        right_hit_box = ss[y:y+h, x:x+w]

        # cv2.imshow('right grey', right_hit_box)
        # cv2.waitKey()

        ret, right_hit_box = cv2.threshold(right_hit_box, 120, 255, cv2.THRESH_BINARY)

        # cv2.imshow('right binary', right_hit_box)
        # cv2.waitKey()

        white_pixels = cv2.countNonZero(right_hit_box)
        # print(f'right wp: {white_pixels}')

        if white_pixels < 2600 and self.right_allowed:
            self.disable_right()
            self.keyboard_client.SendKeys('d')

    def disable_left(self):
        """disable_left

        :return:
        """
        self.left_allowed = False
        thread = Thread(target=self.__disable_left_thread__)
        thread.start()

    def __disable_left_thread__(self):
        """__disable_left_thread__

        :return:
        """
        sleep(.1)
        self.left_allowed = True

    def disable_right(self):
        """disable_right

        :return:
        """
        self.right_allowed = False
        thread = Thread(target=self.__disable_right_thread__)
        thread.start()

    def __disable_right_thread__(self):
        """__disable_right_thread__

        :return:
        """
        sleep(.1)
        self.right_allowed = True

    def disable_top(self):
        """disable_top

        :return:
        """
        self.top_allowed = False
        thread = Thread(target=self.__disable_top_thread__)
        thread.start()

    def __disable_top_thread__(self):
        """__disable_top_thread__

        :return:
        """
        sleep(.1)
        self.top_allowed = True

    def __get_game_window__(self):
        """__get_game_window__

        :return:
        """
        self.game_window = FindWindow(None, GAME_WINDOW_TITLE)
        game_window_rectangle = GetWindowRect(self.game_window)

        top_trim = 450
        sides_trim = 460
        bottom_trim = 90

        self.area_of_interest = (
            game_window_rectangle[0] + sides_trim,
            game_window_rectangle[1] + top_trim,
            game_window_rectangle[2] - sides_trim,
            game_window_rectangle[3] - bottom_trim,
        )

    def run_melee_training(self):
        """run_melee_training

        :return:
        """
        self.melee_training_running = True
        try:
            while True:
                self.__melee_training__()
                if not self.melee_training_running:
                    break
        except KeyboardInterrupt:
            print('Melee Training topped by user!')
        except Exception as exception:
            print(f'Unexpected error occurred: {exception}')

    def toggle_melee_training(self):
        """toggle_melee_training

        :return:
        """
        if self.melee_training_running:
            print('Melee training stopped!')
            self.melee_training_running = False  # run_melee_training loop will stop
        else:
            print('Melee training started!')
            thread = Thread(target=self.run_melee_training)
            thread.start()

    def __enable_key_shortcut__(self):
        """__enable_key_shortcut__

        :return:
        """
        keyboard.on_press_key("f11", lambda _: self.toggle_melee_training())


if __name__ == '__main__':
    ai = SwordsAndSoulsNeverseenAI()
    while True:
        sleep(999999)
