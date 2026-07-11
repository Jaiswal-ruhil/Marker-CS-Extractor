[app]

# (str) Title of your application
title = CS Extractor

# (str) Package name
package.name = csextractor

# (str) Package domain (needed for android packaging)
package.domain = org.ruhiljaiswal

# (str) Source code where the main.py lives
source.dir = .

# (list) Source files to include (let's include py files, png files, txt files)
source.include_exts = py,png,txt

# (list) List of exclusions using pattern matching
#source.exclude_patterns = license,images/*/*.jpg

# (str) Application versioning (method 1)
version = 0.2

# (list) Application requirements
# comma separated e.g. requirements = sqlite3,kivy
requirements = python3,pdfplumber,pandas,openpyxl,reportlab

# (str) Supported orientations (one of landscape, portrait or all)
orientation = landscape

# (bool) Use private storage or not (default is True)
android.private_storage = True

# (int) Android API to use
android.api = 33

# (int) Minimum API required
android.minapi = 21

# (int) Android SDK version to use
android.sdk = 33

# (str) Android NDK version to use
android.ndk = 25b

# (str) Android NDK directory (if empty, it will be automatically downloaded)
#android.ndk_path =

# (str) Android SDK directory (if empty, it will be automatically downloaded)
#android.sdk_path =

# (list) Permissions
android.permissions = READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE

# (list) List of Java .jar files to add to the libs so that py4a can use them
#android.add_jars = foo.jar

# (list) List of Java files to add to the project (e.g. for custom activities)
#android.add_src =

# (list) Android AAR archives to add (must be in libs/ folder)
#android.add_aars =

# (list) Gradle dependencies
#android.gradle_dependencies =

# (str) Android entry point, default is to use start.py of python-for-android
#android.entrypoint = default.平

# (str) Android app theme, default is ok
#android.apptheme = "@android:style/Theme.NoTitleBar"

# (list) Android additionnal libraries to copy into libs/armeabi
#android.add_libs_armeabi = libs/android-v7/libproject.so

# (bool) Logcat filters to use
#android.logcat_filters = *:S python:D

# (str) Android packaging format (apk or aar or aab)
android.release_artifact = apk

# (str) Android architecture to build for (e.g. armeabi-v7a, arm64-v8a)
android.archs = arm64-v8a, armeabi-v7a

# (bool) Copy library instead of making a symlink
#copy_libs = 1

# (list) List of paths to design / asset files
#source.include_folders = assets

[buildozer]

# (int) Log level (0 = error only, 1 = info, 2 = debug (with command output))
log_level = 2

# (int) Display warning if buildozer is run as root (0 = False, 1 = True)
warn_on_root = 1
