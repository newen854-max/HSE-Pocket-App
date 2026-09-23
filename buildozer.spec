[app]

# ------------------------------------------------------------
# Application information
# ------------------------------------------------------------

title = HSE Management System

package.name = hsemanagement

package.domain = com.hse.management

source.dir = .

source.include_exts = py,png,jpg,jpeg,bmp,kv,atlas,ttf,csv,txt

version = 2.0.0


# ------------------------------------------------------------
# Python / Kivy requirements
# ------------------------------------------------------------

requirements = python3,kivy,plyer,xlsxwriter,pillow


# ------------------------------------------------------------
# Android display
# ------------------------------------------------------------

orientation = portrait

fullscreen = 0


# ------------------------------------------------------------
# Android build configuration
# ------------------------------------------------------------

android.api = 35

android.minapi = 23

android.archs = arm64-v8a,armeabi-v7a


# ------------------------------------------------------------
# Android permissions
# ------------------------------------------------------------

android.permissions = READ_MEDIA_IMAGES,READ_MEDIA_VIDEO,READ_MEDIA_VISUAL_USER_SELECTED,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE


# ------------------------------------------------------------
# Android application behavior
# ------------------------------------------------------------

android.allow_backup = True

android.debug_artifact = apk


# ------------------------------------------------------------
# Build logging
# ------------------------------------------------------------

log_level = 2
