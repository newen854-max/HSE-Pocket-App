[app]

# Application name
title = HSE Management System

# Package
package.name = hsemanagement
package.domain = com.hse.management

# Source
source.dir = .
source.include_exts = py,png,jpg,jpeg,bmp,kv,atlas,ttf,csv,txt

# Version
version = 2.0.0

# Python requirements
requirements = python3,kivy,plyer,xlsxwriter,pillow

# Display
orientation = portrait
fullscreen = 0

# ============================================================
# ANDROID
# ============================================================

# Android API
android.api = 35

# Minimum Android API
android.minapi = 23

# Architectures
android.archs = arm64-v8a,armeabi-v7a

# IMPORTANT:
# Automatically accept all Android SDK licenses.
android.accept_sdk_license = True

# Do NOT skip SDK updates/installations.
android.skip_update = False

# Let Buildozer manage its own SDK.
# Leave android.sdk_path empty.

# Backup
android.allow_backup = True

# Permissions
android.permissions = READ_MEDIA_IMAGES,READ_MEDIA_VIDEO,READ_MEDIA_VISUAL_USER_SELECTED,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

# APK
android.debug_artifact = apk

# ============================================================
# BUILD
# ============================================================

log_level = 2
