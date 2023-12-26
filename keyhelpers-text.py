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

REASSIGN_INTERVAL = 180  # seconds


def print_commands_help():
    for name, comb in TRIGGER_KEYS.items():
        print('%22s :' % name.replace('_', " ").capitalize(), " + ".join(comb))



import winclip32

def get_clipboard_text():
    try:
        if winclip32.is_clipboard_format_available("unicode_std_text"):
            return winclip32.get_clipboard_data("unicode_std_text")
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

    yield

    if numlock_on:
        # restore numlock state
        keyboard.send('num lock')

def fix_line_endings(text: str):
    return text.replace('\r\n', '\n')

def copy_selected() -> str:
    # Вставить текст в буфер
    with no_numlock():
        keyboard.send('ctrl + insert')
    sleep(0.02)
    return fix_line_endings(get_clipboard_text())


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

        # Снять выделение, если оно было и прыгнуть в начало слова
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


    # spaces 32 & 160:     

MULTI_SPACES__RE = re.compile(r'[  ]{2,}')
SPACED_PUNCT__RE = re.compile(r'\b[  ]+(?=[.,])')


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
                    send_text_to_editor(changed)

        print(' ...')
        _ = keyboard.stash_state()  # release all keys

    # Wait until all keys of trigger combination are released.
    wait_comb_released(TRIGGER_KEYS.SELECTED_TO_LOWER)
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
    main()
