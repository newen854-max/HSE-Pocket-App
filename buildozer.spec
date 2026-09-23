[app]

title = HSE Management System
package.name = hsemanagement
package.domain = com.hse.management

source.dir = .
source.include_exts = py,png,jpg,jpeg,bmp,kv,atlas,ttf,csv,txt

version = 2.0.0

# Keep Android dependencies minimal.
# xlsxwriter is pure Python and can be added later if needed.
# Pillow is deliberately removed from the Android build.
requirements = python3,kivy,plyer,xlsxwriter

orientation = portrait
fullscreen = 0

# ------------------------------------------------------------
# Android
# ------------------------------------------------------------

android.api = 35
android.minapi = 24
android.ndk = 28c

android.accept_sdk_license = True

android.archs = arm64-v8a

android.copy_libs = True

android.allow_backup = True

android.debug_artifact = apk

android.permissions = INTERNET

# ------------------------------------------------------------
# Python-for-Android
# ------------------------------------------------------------

p4a.bootstrap = sdl2
p4a.branch = master

# ------------------------------------------------------------
# Build
# ------------------------------------------------------------

log_level = 2