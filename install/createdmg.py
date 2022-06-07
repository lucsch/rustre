#!/usr/bin/python

import argparse

import os
import subprocess


def create_installer(mod_name, app_path, output_path, use_version=True):
    version = 0
    basepath = os.path.join(os.path.dirname(__file__), "..")
    if use_version is True:
        with open(os.path.join(basepath, mod_name, "version.py")) as reader:
            for line in reader:
                if line.startswith("COMMIT_NUMBER"):
                    version = int(line[17:-2])

        if version:
            print("Git version is:", version)
        else:
            print("Error getting Git version number!")

    mycommand = [
        'hdiutil',
        'create',
        '-volname', mod_name,
        '-srcfolder', app_path,
        os.path.join(output_path, mod_name + "_" + str(version) + ".dmg")]
    print(mycommand)

    try:
        p = subprocess.Popen(mycommand)
        p.wait()
    except (OSError, ValueError):
        print("Error creating DMG!")
        return
    print("Creating DMG finished!")


def delete_binary(app_path):
    for name in os.listdir(app_path):
        if os.path.isfile(os.path.join(app_path, name)) and not name.startswith("."):
            print("removing : " + name)
            os.remove(os.path.join(app_path, name))


##########################################################
# Main function, parse command line arguments
##########################################################
if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('mod_name', help='Module and Installer Name')
    parser.add_argument('app_path', help='Application Path')
    parser.add_argument('output_path', help='Output Path')
    args = parser.parse_args()
    delete_binary(args.app_path)
    create_installer(args.mod_name, args.app_path, args.output_path, use_version=True)
