#!/usr/bin/python

##########################################################
# Create a APP package for OSX
# (c) Lucien SCHREIBER 2020
# usage :
#     createapp.py
##########################################################
import os
import sys
import subprocess
import fileinput
import argparse

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from rustre.createversion import GitVersion  # noqa: E402

ACTIVE_PLATEFORM = ["Windows", "Linux", "OSX"]


class CreateApp(object):
    def __init__(self, plateform="OSX"):
        if plateform not in ACTIVE_PLATEFORM:
            raise ValueError("plateform must be one of %r." % ACTIVE_PLATEFORM)

        self.basepath = os.path.join(os.path.dirname(__file__), "..")
        self.binpath = os.path.join(self.basepath, "bin")
        self.plateform = plateform
        self.iconfile = os.path.join(self.basepath, "art", self._get_icon())
        self.m_commit_number = ""

    def _get_icon(self):
        """return the icon based on the plateform"""
        icon = "rustre.icns"
        if self.plateform == ACTIVE_PLATEFORM[0]:  # Windows
            icon = "rustre.ico"
        elif self.plateform == ACTIVE_PLATEFORM[1]:  # Linux
            icon = "rustre.png"
        return icon

    def update_version(self):
        """update the about.py file with the version number"""
        my_version = GitVersion()
        my_version.write_to_file(os.path.join(self.basepath, "rustre", "version.py"))
        self.m_commit_number = my_version.m_commit_number
        return True

    def modify_spec_file(self):
        """modifiy the spec file before building"""
        if self.plateform == ACTIVE_PLATEFORM[2]:  # OSX
            my_spec_path = os.path.join(self.binpath, "rustre_{}.spec".format(self.m_commit_number))
            my_spec_path = os.path.normpath(my_spec_path)
            for line in fileinput.input(my_spec_path, inplace=1):
                if "bundle_identifier=None)" in line:
                    print("             bundle_identifier=None,")
                    print("             info_plist={")
                    print("                 'CFBundleShortVersionString': '1.0.{}',".format(self.m_commit_number))
                    print("                 'NSHumanReadableCopyright': '(c) 2020, Lucien SCHREIBER',")
                    print("                 'NSHighResolutionCapable': 'True',")
                    print("                 'NSRequiresAquaSystemAppearance': 'No'")
                    print("             })")
                else:
                    print(line[:-1])

    def create_exe(self):
        """run pyInstaller to create the exe"""
        if not os.path.exists(self.binpath):
            os.makedirs(self.binpath)

        command = [
            "pyi-makespec",
            "--onefile",
            "--windowed",
            "--hidden-import=wx",
            "--hidden-import=wx._xml",
            "--hidden-import=pkg_resources.py2_warn",
            # "--hidden-import=sqlalchemy.ext.baked",
            "-nrustre_{}".format(self.m_commit_number),
            "--icon={}".format(self.iconfile),
            # "--add-data={}".format(os.path.join(self.basepath, "rustre", "html") + os.pathsep + "html"),
            os.path.join(self.basepath, "rustre", "__main__.py")]
        print(command)
        try:
            p = subprocess.Popen(command, cwd=self.binpath)
            p.wait()
        except (OSError, ValueError):
            print("Error running: " + " ".join(command))
            return False

        self.modify_spec_file()

        # run pyinstaller with spec file
        try:
            p = subprocess.Popen(["pyinstaller", "rustre_{}.spec".format(self.m_commit_number), "-y"],
                                 cwd=self.binpath)
            p.wait()
        except (OSError, ValueError):
            print("Error running : pyinstaller rustre.spec")
            return False


##########################################################
# Main function, parse command line arguments
##########################################################
if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('plateform', help="choose a plateform. Supported values are : " + ", ".join(ACTIVE_PLATEFORM))
    args = parser.parse_args()

    myApp = CreateApp(plateform=args.plateform)
    myApp.update_version()
    myApp.create_exe()
