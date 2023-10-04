import copy
import click
import pathlib

import move
import translate

import windows
import utilities


if __name__ == "__main__":

    print()
    directory = click.prompt("Directory", type=click.Path(exists=True, file_okay=False, dir_okay=True))
    processes = click.prompt("Processes", type=click.IntRange(min=1, max=16), default=8)
    full = click.prompt("Full", type=click.BOOL, default=False)
    safe = click.prompt("Safe", type=click.BOOL, default=True)
    print()

    directory_src = pathlib.Path(f"{directory}")
    directory_dst = pathlib.Path(f"{directory} - Translated")

    watch = utilities.Watch()

    windows.clear_temp()
    move.move(full, directory_src, directory_dst, copy.deepcopy(watch))
    translate.translate(safe, directory_dst, processes, copy.deepcopy(watch))
    translate.translate(safe, directory_dst, 1, copy.deepcopy(watch))

    print()
    click.pause("Press any key to exit...")
    print()
