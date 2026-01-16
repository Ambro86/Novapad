#----------------------------------------------------------------
# Generated CMake target import file for configuration "Release".
#----------------------------------------------------------------

# Commands may need to know the format version.
set(CMAKE_IMPORT_FILE_VERSION 1)

# Import target "nghttp2::nghttp2_static" for configuration "Release"
set_property(TARGET nghttp2::nghttp2_static APPEND PROPERTY IMPORTED_CONFIGURATIONS RELEASE)
set_target_properties(nghttp2::nghttp2_static PROPERTIES
  IMPORTED_LINK_INTERFACE_LANGUAGES_RELEASE "C"
  IMPORTED_LOCATION_RELEASE "${_IMPORT_PREFIX}/lib/nghttp2.lib"
  )

list(APPEND _cmake_import_check_targets nghttp2::nghttp2_static )
list(APPEND _cmake_import_check_files_for_nghttp2::nghttp2_static "${_IMPORT_PREFIX}/lib/nghttp2.lib" )

# Commands beyond this point should not need to know the version.
set(CMAKE_IMPORT_FILE_VERSION)
