""" Новые горячие клавиши, которые будут помогать работать с текстом.

Первоначальная идея - использовать для корректировки текста, набранного голосом при помощи диктовки, когда возможностей Punto Switcher уже не хватает.


Функции, который реализованы:
  — Изменение регистра первой буквы слова (активируется по отпусканию Win+Alt+Up) (требуется проверка на Windows 10+ – не является ли это сочетание системным).
  — Удаление из текста дублирующихся пробелов, а также пробелов перед точками и запятыми.
  — Перевод всего текста в нижний регистр.

  (Используемые сочетания клавиш перечислены после основных импортов и могут быть изменены.)

Дополнительная вещь, которая сделана, — это восстановление содержимого буфера обмена после выполнения команды.

Особенности.

Зависимости: keyboard, win32api, winclip32
Работает только на Windows.
Работает в фоне, но при этом окно консоли питона остаётся на панели задач (чтобы его скрывать, я применяю MinimizeToTray).
Экстренный выход из программы - по сочетанию клавиш Win + Esc.

Заметки по реализации.

Раз в минуту горячие клавиши переназначаются (это может быть полезно, если другие программы перехватывают ввод).

Работа с буфером обмена была проблемной, пока я использовал Control + C / Control + V.
Текст копировался плохо, то есть ненадёжно: иногда копировался, а иногда менялась громкость вместо этого.
Затем я переключился на использование Control + Insert / Shift + Insert, после чего всё работает более стабильно.
"""


from contextlib import contextmanager
from math import copysign
import re
import os
from time import sleep

from win32api import GetKeyState
from win32con import VK_NUMLOCK

import keyboard

class adict(dict):
    'imitation of `adict` from PyPI'
    __getattr__ = dict.__getitem__
    __setattr__ = dict.__setitem__


TRIGGER_KEYS = adict()

# keys for commands to trigger
TRIGGER_KEYS.TOGGLE_CASE = ['alt', 'windows', 'up']
# TRIGGER_KEYS.TOGGLE_CASE = ['windows', 'F4']
# TRIGGER_KEYS.SELECT_WORD = ['control', 'windows', 'd']  # similar to Ctrl+Shift+Space or Ctrl+D (once) in Sublime - triggers new Desktop on Windows 10!!!
TRIGGER_KEYS.SHRINK_MULTIPLE_SPACES = ['left shift', 'windows', 'right shift']
# TRIGGER_KEYS.SHRINK_MULTIPLE_SPACES = ['alt', 'c']
TRIGGER_KEYS.SELECTED_TO_LOWER = ['shift', 'windows', 'left']

TRIGGER_KEYS.CUSTOM_TRANSFORM = ['shift', 'windows', 'alt']


REASSIGN_INTERVAL = 180  # seconds

POS_CORRECTION = True  # alg will try to sync tracking position with actual position where possible
MAX_CORRECTED_DISTANCE = 10  # how far actual position can be (in chars)


TM = None # TextManipulator

def print_commands_help():
    for name, comb in TRIGGER_KEYS.items():
        print('%22s :' % name.replace('_', " ").capitalize(), " + ".join(comb))



import winclip32

def get_clipboard_text():
    try:
        if winclip32.is_clipboard_format_available("unicode_std_text"):
            content = winclip32.get_clipboard_data("unicode_std_text")
            return fix_line_endings(content)
            # return content

    except winclip32.errors.clipboard_format_is_not_available:
        pass
    return ''

def set_clipboard_text(t):
    winclip32.set_clipboard_data("unicode_std_text", t)


@contextmanager
def backup_of_clipboard():
    backup = get_clipboard_text()
    # wait some time to complete all internal actions (?)
    # sleep(0.09)
    try:
        yield
    finally:
        # Это очистит буфер обмена, если до этого в нём был не текст
        set_clipboard_text(backup)


def wait_comb_released(keys, check_interval=0.03):
    """ Wait until all the keys are released."""
    while True:
        if any(map(keyboard.is_pressed, keys)):
            sleep(check_interval)
        else:
            break


@contextmanager
def no_numlock(verbose=False):
    """temporarily turn numlock OFF while doing wrapped command, then enable back if it was enabled before.
    Simulating shift + something else does not work when numlock is enabled (!)"""
    numlock_on = GetKeyState(VK_NUMLOCK)

    if numlock_on:
        if verbose: print("Num Lock is on")
        # temporarily turn numlock OFF
        keyboard.send('num lock')

    try: yield
    except: raise
    finally:
        if numlock_on:
            # restore numlock state
            keyboard.send('num lock')

def fix_line_endings(text: str):
    return text.replace('\r\n', '\n')

def copy_selected() -> str:
    # Взять текст в буфер
    with no_numlock():
        keyboard.send('ctrl + insert')
    sleep(0.02)
    return get_clipboard_text()


def paste_text(text):
    # Вставить текст через буфер
    set_clipboard_text(text)
    sleep(0.02)
    with no_numlock():
        keyboard.send('shift + insert')


def send_text_to_editor(text, len_treshold=100):
    """ "Smart" Ctrl+V: depending on text size, use typewrite or direct paste via Shift + Insert."""
    if len(text) < len_treshold:
        # Визуально наглядно напечатает текст, заменяя старый (удобно, если текст не длинный)
        keyboard.write(text)
    else:
        # Вставит текст через буфер, заменяя старый (быстрее, если текст длинный)
        paste_text(text)


class MutableString(object):
    def __init__(self, data):
        self.data = list(data)
    def __repr__(self):
        return "".join(self.data)
    __str__ = __repr__
    def __setitem__(self, index, value):
        if type(index) == slice:
            self.data[index] = list(value)
        else:
            self.data[index] = value
    def __getitem__(self, index):
        if type(index) == slice:
            return "".join(self.data[index])
        return self.data[index]
    def __delitem__(self, index):
        del self.data[index]
    def __add__(self, other):
        self.data.extend(list(other))
    def __len__(self):
        return len(self.data)

def findall(sub, string):
    """
    @see: https://stackoverflow.com/a/3874760/12824563
    >>> text = "Allowed Hello Hollow"
    >>> tuple(findall('ll', text))
    (1, 10, 16)
    """
    index = -len(sub)
    try:
        while True:
            index = string.index(sub, index + len(sub))
            yield index
    except ValueError:
        pass


class TextManipulator:
    def __init__(self):
        self.pos = 0
        self.sel = [0, 0]  # left pos, right pos of selection
        self.sel_edge = 0  # 0 or 1 - which side of selection the cursor is at
        self.mirror = ""
        self._pos_shift = 0
        self._validation_tries = 0  # limit: say, 10 or 100
    clear = __init__

    def set_string(self, string, selected_in_editor=False, relative_pos=None):
        self.mirror = MutableString(string)
        if selected_in_editor:
            self.sel = [0, len(string)]
        if relative_pos is not None:
            self.pos = relative_pos

    def get_string(self):
        return str(self.mirror)

    def sel_len(self):
        return self.sel[1] - self.sel[0]

    def move_to(self, pos, shift=False):
        # when not selecting, no prior selection is allowed (limitation of current implementation: extra checks needed).
        self.move_by(pos - self.pos, shift=shift)

    def move_by(self, steps, shift=False):
        """Move cursor left (negative) or right (positive) optionally holding Shift."""
        if steps == 0: return
        n_steps = int(abs(steps))
        steps = int(steps)
        is_right = steps > 0

        keys = 'right' if is_right else 'left'  # arrow
        if shift: keys = 'shift + ' + keys

        for _ in range(n_steps):
            ###
            # print('   move_by ::', keys)
            # sleep(.25)
            ###
            keyboard.send(keys)

        if shift:
            self.pos += steps
            if self.sel_len() <= n_steps:
                # flip selection
                self.sel[:] = self.sel[::-1]
                self.sel_edge = is_right
            self.sel[self.sel_edge] = self.pos
        else:
            if self.sel_len() > 0:
                # no more selection (we've moved to the edge first)
                self.pos = self.sel[is_right] + int(copysign(n_steps - 1, steps))
            else:
                self.pos += steps
            self.sel = [self.pos, self.pos]

    def select_range(self, start, stop):
        from_left  = abs(self.pos - start)
        from_right = abs(stop - self.pos)
        if from_right < from_left:
            # select from right
            self.move_to(stop)
            self.move_to(start, shift=True)
        else:
            # select from left
            self.move_to(start)
            self.move_to(stop, shift=True)

    def delete_selected_text(self):
        '& update self.sel'
        keyboard.send('delete')
        # track changes internally
        self.mirror[self.sel[0] : self.sel[1]] = ''
        # update sel positions
        self.sel[1] = self.sel[0]
        # update pos
        self.pos = self.sel[0]

    def write_text(self, text):
        if not text:
            return
        send_text_to_editor(text)
        self.mirror[self.pos : self.pos] = text  # track changes
        self.pos += len(text)

    def validate_position(self):
        """ Using and extending selection if required, verify current position (self.pos) and amend it if needed and possible."""

        selected_extra = 0
        sel_direction = +1  # -1 (left) or +1 (right)

        def finalize():
            # undo selection changes
            self.move_by(-sel_direction * selected_extra, shift=True)

        def apply_correction(correction: int):
            self.pos += correction
            self.sel[0] += correction
            self.sel[1] += correction
            print('Appiled correction to pos: %+d' % correction)


        # if self.sel_len() > 0:
        #     self.move_by(-1)  # reset sel to left

        if self.sel_len() <= 0:
            # make selection
            if self.pos + (MAX_CORRECTED_DISTANCE // 2) > len(self.mirror):
                sel_direction = -1  # grow to left
            self.move_by(sel_direction, shift=True)
            selected_extra += 1
            ###
            print(f'### validate_position: selected 1 char to {sel_direction}')


        # copy selection to check where we are
        selected = copy_selected()
        tracked_sel = self.mirror[slice(*self.sel)]
        if len(selected) != self.sel_len():
            # stop working on serious error, ex. if user has clicked on different place in editor or switched an application
            raise RuntimeError(f'Cannot validate position via selection! (selected `{selected}`, expected: `{tracked_sel}`).')

        if selected == tracked_sel:
            ###
            print(f'### validate_position: OK!')

            # OK!
            finalize()
            return True

        # try to correct self.pos...

        mirror_str = str(self.mirror)
        if selected not in mirror_str:
            # The issue may be also due moving out of the edge of processed piece of text (reflected by self.mirror).
            # stop working on serious error, ex. if user has clicked on different place in editor or switched an application
            raise RuntimeError(f'Cannot validate position via selection! (cannot find selected `{selected}`; expected: `{tracked_sel}`).')

        ### print(f'### validate_position: mirror_str: `{mirror_str}`\n')

        # find all occurences
        occurences = list(findall(selected, mirror_str))

        ###
        print(f'### validate_position: occurences: {occurences}')

        # find closest occurence
        pos_of_sel = self.sel[0]

        if len(occurences) == 1:  # and abs(occurences[0] - pos_of_sel) <= MAX_CORRECTED_DISTANCE:
            # this is probably the only candidate to move to
            # do the correction
            correction = int(occurences[0] - pos_of_sel)
            apply_correction(correction)

            finalize()
            # we have corrected it now, but still need to double-check it
            return False

        # find nearest occurence
        dst_to_occurences = sorted([
                    (abs(p - pos_of_sel), p)
                    for p in occurences
                    if abs(p - pos_of_sel) <= MAX_CORRECTED_DISTANCE
                ])


        ###
        print(f'### validate_position: dst_to_occurences: {dst_to_occurences}')


        if len(dst_to_occurences) == 1:
            # this is probably the only candidate to move to
            # do the correction
            correction = int(dst_to_occurences[0][1] - pos_of_sel)
            apply_correction(correction)

            finalize()
            # we have corrected it now, but still need to double-check it
            return False

        else:
        # elif 0:
            # there's need to clarify which one we should stick to.
            # grow selection until we have only one candidate.
            ...
            grow_times = min(
                    MAX_CORRECTED_DISTANCE - self.sel_len(), # max selection size is MAX_CORRECTED_DISTANCE
                    len(self.mirror) - self.pos  # the rest of mirror
            )
            ###
            print(f'### validate_position: grow_times: {grow_times}')

            for _i in range(grow_times):

                print('trying to recover position, take', _i + 1)

                # grow selection
                self.move_by(sel_direction)
                selected_extra += 1

                # (repeat some things from above ...)
                selected = copy_selected()
                occurences = list(findall(selected, mirror_str))
                if not occurences:
                    break
                dst_to_occurences = sorted([
                            (abs(p - pos_of_sel), p)
                            for p in occurences
                            if abs(p - pos_of_sel) <= MAX_CORRECTED_DISTANCE
                        ])
                if not dst_to_occurences:
                    break

                if len(dst_to_occurences) == 1:
                    # we've found the most probable candidate.
                    # do the correction
                    correction = int(dst_to_occurences[0][1] - pos_of_sel)
                    apply_correction(correction)

                    finalize()
                    # we have corrected it now, but still need to double-check it
                    return False

        finalize()
        raise RuntimeError(f'Cannot recover from unexpected position change, stopping. Selected: `{selected}`, occurences: {occurences}')



    def apply_transformation(self, t):
        ### print('self._pos_shift:', self._pos_shift)
        start, stop = t.position.start + self._pos_shift, t.position.stop + self._pos_shift

        if start == stop:
            # insert text
            if not t.replacement: return

            self.move_to(start)
            # send_text_to_editor(t.replacement)
            # self.mirror[start:stop] = t.replacement
            # self._pos_shift += len(t.replacement)

        else:
            # select target area
            self.select_range(start, stop)

            if POS_CORRECTION:
                # TODO: verify that target area is selected correctly (avoiding unexpected shifts)
                for _ in range(10):
                    # try several times in case that user keeps changing position manually
                    if self.validate_position():
                        break
                    ###
                    print('### correction loop: try again...')
                else:
                    raise RuntimeError('Cannot recover invalid position after several tries, stopping.')
                # ...
                # self.select_range(start, stop)

            self.delete_selected_text()

        self.write_text(t.replacement)  # + track changes internally
        self._pos_shift += len(t.replacement) - (stop - start)


    def apply_transformations(self, transformations):
        """
            transformations: iterable of adicts:
            [
                {position: slice(3,6), replacement: "abc"},  # replace
                {position: slice(6,7), replacement: ""},     # remove
                {position: slice(7,7), replacement: "new"},  # insert
            ]

            The order must be ascending of position; no overlaps.
        """
        if transformations:  # anything to process

            # go to leftmost replacement pos ...
            leftmost_pos = transformations[0].position

            if (sl := self.sel_len()) > 0:
                # decide which direction to go (we can go both sides at first step)
                from_left  = abs(self.sel[0] - leftmost_pos.start)
                from_right = abs(leftmost_pos.stop - self.sel[1])
                if from_right < from_left:
                    # move to right side of leftmost replacement
                    self.move_by(+1)
                    # self.move_by(-from_right)
                    ### print('remove selection to its right side')
                else:
                    # move to left side of leftmost replacement
                    self.move_by(-1)
                    # self.move_by(+from_left)
                    ### print('remove selection to its left side')

            self._pos_shift = 0
            for t in transformations:
                ###
                print(t, '...')
                # sleep(.1)
                ###
                self.apply_transformation(t)


        if self.pos < len(self.mirror):
            # skip the rest of string
            self.move_to(len(self.mirror))


def get_transformations_for_coherent_strings(a, b):
    if len(a) != len(b):
        raise ValueError('strings must be of equal lengths!')

    transformations = []
    extending_last = False

    for i, (x, y) in enumerate(zip(a, b)):
        if x == y:
            extending_last = False
            continue

        if not extending_last:
            last = adict(position = slice(i, i + 1), replacement = y)
            transformations.append(last)
            extending_last = True
        else:
            last = transformations[-1]
            # extend range and replacement
            last.position = slice(last.position.start, i + 1)
            last.replacement += y

    return transformations


####### ---------------------------------------------------- ########
####### Main helper functions triggered by global keystrokes ########
####### ---------------------------------------------------- ########


def toggle_word_title(*args, **kw):
    """select a word to its beginning, copy, change case of the first letter and paste back."""

    print('hotkey is triggered.')

    def go(*_):
        print(' - Go !  --  toggle Case')

        numlock_on = GetKeyState(VK_NUMLOCK)

        if numlock_on:
            print("Num Lock is on")
            # temporarily turn numlock OFF
            keyboard.send('num lock')
            # simulating shift + arrows does not work when numlock is enabled (!)

        selected_before = get_clipboard_text()

        # Снять выделение, если оно было
        keyboard.send('right, left')

        # select a word to its beginning
        keyboard.send('shift + ctrl + left')

        if numlock_on:
            # restore numlock state
            keyboard.send('num lock')

        # if True:
        with backup_of_clipboard():

            ###
            # print('Sending Ctrl+C')
            # sleep(1)

            # copy
            # keyboard.send('ctrl + c')
            keyboard.send('ctrl + insert')


            ###
            # print('Sending Ctrl+C DONE.')
            # sleep(1)


            sleep(0.1)
            selected = get_clipboard_text()

            if not selected:
                print('-- (nothing copied.)')
                # pass
            if selected == selected_before:
                # not working with text field now
                # return
                print('-- (copied the same.?)')
                # pass
            else:
                print('Got: `%s`' % selected)
                changed = selected.lower()
                if changed == selected:
                    changed = selected.capitalize()

                if changed != selected:
                    print('Put: `%s`' % changed)
                    send_text_to_editor(changed)

                    ###
                    print(get_transformations_for_coherent_strings(selected, changed))

                    # paste back
                    # set_clipboard_text(changed)
                    # keyboard.send('ctrl + v')

                    # keyboard.write(changed)

        print(' ...')
        _ = keyboard.stash_state()  # release all keys

    # Wait until all keys of trigger combination are released.
    wait_comb_released(TRIGGER_KEYS.TOGGLE_CASE)
    keyboard.call_later(go, delay=0.0)


def select_word(*args, **kw):
    """Jump to word's beginning and select its end. Same as double-click in most of editors."""

    print('hotkey is triggered.')

    def go(*_):
        print(' - Go ! -- Select a WORD')

        numlock_on = GetKeyState(VK_NUMLOCK)

        if numlock_on:
            print("Num Lock is on")
            # temporarily turn numlock OFF
            keyboard.send('num lock')
            # simulating shift + arrows does not work when numlock is enabled (!)

        selected_before = get_clipboard_text()

        # Снять выделение, если оно было, и прыгнуть в начало слова
        keyboard.send('right, ctrl + left')

        # select a word to its beginning
        keyboard.send('shift + ctrl + right')

        if numlock_on:
            # restore numlock state
            keyboard.send('num lock')

        print(' ...')

        _ = keyboard.stash_state()  # release all keys

    # Wait until all keys of trigger combination are released.
    wait_comb_released(TRIGGER_KEYS.SELECT_WORD)
    keyboard.call_later(go, delay=0.0)


    # spaces 32 & 160:  '    '

MULTI_SPACES__RE = re.compile(r'[  ]{2,}')
SPACED_PUNCT__RE = re.compile(r'\b[  ]+(?=[.,)])|(?<=\()[  ]+(?=\w)')


def shrink_multiple_spaces(*args, **kw):
    """Within selected, remove double and more subsequent spaces."""

    print('hotkey is triggered.')

    def go(*_):
        print(' - Go !  --  shrink multiple spaces')

        # numlock_on = GetKeyState(VK_NUMLOCK)

        # if numlock_on:
        #     print("Num Lock is on")
        #     # temporarily turn numlock OFF
        #     keyboard.send('num lock')
        #     # simulating shift + arrows does not work when numlock is enabled (!)

        # selected_before = get_clipboard_text()

        # # Снять выделение, если оно было
        # keyboard.send('right, left')

        # # select a word to its beginning
        # keyboard.send('shift + ctrl + left')

        # if numlock_on:
        #     # restore numlock state
        #     keyboard.send('num lock')

        # if True:
        with backup_of_clipboard():

            ###
            # print('Sending Ctrl+C')
            # sleep(1)

            # copy
            # keyboard.send('ctrl + c')
            keyboard.send('ctrl + insert')


            ###
            # print('Sending Ctrl+C DONE.')
            # sleep(1)


            sleep(0.05)
            selected = get_clipboard_text()

            if not selected:
                print('-- (nothing copied.)')
                # pass
            # if selected == selected_before:
            #     # not working with text field now
            #     # return
            #     print('-- (copied the same.?)')
            #     # pass
            else:
                print('Got: `%s`' % selected)
                # Убираем лишние пробелы
                changed = MULTI_SPACES__RE.sub(" ", selected)
                changed = SPACED_PUNCT__RE.sub('', changed)

                if changed == selected:
                    print('No changes made.')
                    # changed = selected.capitalize()

                if changed != selected:
                    # paste back
                    print('Put: `%s`' % changed)
                    # paste_text(changed)
                    send_text_to_editor(changed)

                    # set_clipboard_text(changed)
                    # sleep(0.1)
                    # with no_numlock():
                    #     keyboard.send('shift + insert')

                    # # print('Put: `%s`' % changed)
                    # # keyboard.write(changed)

        print(' ...')
        _ = keyboard.stash_state()  # release all keys

    # Wait until all keys of trigger combination are released.
    wait_comb_released(TRIGGER_KEYS.SHRINK_MULTIPLE_SPACES)
    keyboard.call_later(go, delay=0.0)


def shrink_multiple_spaces_advanced(*args, **kw):
    """Within selected, remove double and more subsequent spaces, updating only required areas within the text."""

    print('hotkey is triggered.')

    def go(*_):
        print(' - Go !  --  shrink multiple spaces (Advanced)')

        # if True:
        with backup_of_clipboard():
          with no_numlock():

            selected = copy_selected()
            if not selected:
                print('-- (nothing copied.)')
                return

            # verify the selection ...
            if 1:
                # step_left
                keyboard.send('shift + left')
                selected_m1 = copy_selected()
                keyboard.send('shift + right')  # move back

                if not (1 <= abs(len(selected) - len(selected_m1)) <= 2) or (selected_m1 not in selected and selected not in selected_m1):
                    print('-- (selection does not seem to be valid, skipping.)')
                    print(f'-- ({len(selected_m1)} !<! {len(selected)})')
                    print(f'-- (`{selected_m1}` !<! `{selected}`)')
                    return


            print('Got: `%s`' % selected)

            selection = selected

            # Убираем лишние пробелы ...

            re_to_replacement = {
                MULTI_SPACES__RE: " ",
                SPACED_PUNCT__RE: '',  # later re rewrites entries of earlier re
            }

            pos2replacement_pair = dict(sorted(
                (
                    (
                        m.start(),
                        (m[0], repl)  # (needle=m[0], replacement=repl)
                    )
                       for regex, repl in re_to_replacement.items()
                       for m in regex.finditer(selection)
                ),
                key=lambda t: t[0]  # sort by pos only, not by match
            ))
            if not pos2replacement_pair:
                print('-- No changes required.')
                # go to the right edge of selection
                keyboard.send('right')
                return

            # ## print(pos2replacement_pair)

            # go to the left edge of selection
            keyboard.send('left')

            # run over string and change/replace required chars only

            pos_in_selection = 0

            for pos, (needle, replacement) in pos2replacement_pair.items():
                chars_to_skip = pos - pos_in_selection
                for _ in range(chars_to_skip):
                    keyboard.send('right')
                pos_in_selection += chars_to_skip

                if needle.endswith(replacement) or not replacement:
                    # remove few leading chars
                    leading_chars_to_del = len(needle) - len(replacement)
                    for _ in range(leading_chars_to_del):
                        keyboard.send('delete')
                    for _ in range(len(replacement)):
                        keyboard.send('right')
                else:
                    # replace chars
                    leading_chars_to_del = len(needle) - len(replacement)
                    for _ in range(len(needle)):
                        keyboard.send('delete')
                    send_text_to_editor(replacement)

                pos_in_selection += len(needle)

            # skip the rest of string
            chars_to_skip = len(selection) - pos_in_selection
            for _ in range(chars_to_skip):
                keyboard.send('right')



        print(' ...')
        _ = keyboard.stash_state()  # release all keys

    # Wait until all keys of trigger combination are released.
    wait_comb_released(TRIGGER_KEYS.SHRINK_MULTIPLE_SPACES)
    keyboard.call_later(go, delay=0.0)


def selected_to_lower(*args, **kw):
    """Turn all selected characters to lower-case."""

    print('hotkey is triggered.')

    def go(*_):
        print(' - Go !  --  Selected to lower')

        # if True:
        with backup_of_clipboard():

            # copy
            # keyboard.send('ctrl + c')
            keyboard.send('ctrl + insert')

            sleep(0.1)
            selected = get_clipboard_text()
            # selected = copy_selected()

            if not selected:
                print('-- (nothing copied.)')
                # pass
            # if selected == selected_before:
            #     # not working with text field now
            #     # return
            #     print('-- (copied the same.?)')
            else:
                print('Got: `%s`' % selected)
                changed = selected.lower()
                if changed == selected:
                    print('Everything is already in lowercase.')
                    # changed = selected.upper()

                if changed != selected:
                    print('Put: `%s`' % changed)
                    # paste_text(changed)

                    ###
                    TM.set_string(selected, selected_in_editor=True)
                    with no_numlock():
                        TM.apply_transformations(get_transformations_for_coherent_strings(selected, changed))

                    # send_text_to_editor(changed)

        print(' ...')
        _ = keyboard.stash_state()  # release all keys

    # Wait until all keys of trigger combination are released.
    wait_comb_released(TRIGGER_KEYS.SELECTED_TO_LOWER)
    keyboard.call_later(go, delay=0.0)


def custom_transform(*args, **kw):
    """ Do some custom transformation in selected text.
    """

    print('custom hotkey is triggered.')

    def go(*_):
        print(' - Go !  --  Custom transform.')

        # if True:
        with backup_of_clipboard():

            # copy
            # keyboard.send('ctrl + c')
            keyboard.send('ctrl + insert')

            sleep(0.1)
            selected = get_clipboard_text()
            # selected = copy_selected()

            if not selected:
                print('-- (nothing copied.)')
            else:
                print('Got: `%s`' % selected)
                # ==============================
                ## sequence.action_kind = 'sequence'
                ## >>>
                ## set_enum_value(sequence, action_kind, 'sequence')
                # changed = re.sub(
                #     r'(\w+)\.(\w+) = (["\']\w+["\']|\w+);?',
                #     r'set_enum_value(\1, \2, \3)',
                #     selected
                # )

                # transform HTML: remove style attribute, replace nbsp.
                changed = selected.replace("&nbsp;", " ")
                changed = changed.replace(' style="border-color: var(--border-color);"', '')
                # ==============================

                if changed == selected:
                    print('Everything is already updated.')

                if changed != selected:
                    print('Put: `%s`' % changed)
                    # paste_text(changed)
                    keyboard.send('delete')
                    keyboard.write(changed)

                    # ###
                    # TM.set_string(selected, selected_in_editor=True)
                    # with no_numlock():
                    #     TM.apply_transformations(get_transformations_for_coherent_strings(selected, changed))


        print(' ...')
        _ = keyboard.stash_state()  # release all keys

    # Wait until all keys of trigger combination are released.
    wait_comb_released(TRIGGER_KEYS.CUSTOM_TRANSFORM)
    keyboard.call_later(go, delay=0.0)


def main():

    def assign_hotkeys():
        # see '_winkeyboard.py' for keynames

        keyboard.add_hotkey(
            ' + '.join(TRIGGER_KEYS.TOGGLE_CASE),
            toggle_word_title,
            suppress=False,
            trigger_on_release=False,
            timeout=1,
        )
        # keyboard.add_hotkey(
        #     ' + '.join(TRIGGER_KEYS.SELECT_WORD),
        #     select_word,
        #     suppress=False,
        #     trigger_on_release=False,
        #     timeout=1,
        # )
        keyboard.add_hotkey(
            ' + '.join(TRIGGER_KEYS.SHRINK_MULTIPLE_SPACES),
            shrink_multiple_spaces_advanced,
            suppress=False,
            trigger_on_release=False,
            timeout=1,
        )
        keyboard.add_hotkey(
            ' + '.join(TRIGGER_KEYS.SELECTED_TO_LOWER),
            selected_to_lower,
            suppress=False,
            trigger_on_release=False,
            timeout=1,
        )
        keyboard.add_hotkey(
            ' + '.join(TRIGGER_KEYS.CUSTOM_TRANSFORM),
            custom_transform,
            suppress=False,
            trigger_on_release=False,
            timeout=1,
        )

        # keyboard.add_hotkey(mod+' + F5',  simulate, args=(mod, 'alt + break'.split(', ')), suppress=True)

        os.system('CLS')
        print_commands_help()
        print('Ready...')

    def reassign_hotkeys():
        keyboard.clear_all_hotkeys()
        assign_hotkeys()

        # Quit combination
        keyboard.register_hotkey('windows + esc', suppress=False, callback=full_quit)

        keyboard.call_later(reassign_hotkeys, delay=REASSIGN_INTERVAL)
        
    # Start reassigning loop
    assign_hotkeys()
    keyboard.call_later(reassign_hotkeys, delay=REASSIGN_INTERVAL)


    def full_quit(*_):
        keyboard.clear_all_hotkeys()
        _ = keyboard.stash_state()  # release all keys
        print('Bye.')
        os._exit(0)


    # Quit combination
    keyboard.register_hotkey('windows + esc', suppress=False, callback=full_quit)
    # keyboard.wait('right windows + esc')


if __name__ == '__main__':

    TM = TextManipulator()
    main()
