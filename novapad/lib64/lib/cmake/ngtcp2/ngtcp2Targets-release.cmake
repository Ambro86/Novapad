#----------------------------------------------------------------
# Generated CMake target import file for configuration "Release".
#----------------------------------------------------------------

# Commands may need to know the format version.
set(CMAKE_IMPORT_FILE_VERSION 1)

# Import target "ngtcp2::ngtcp2_static" for configuration "Release"
set_property(TARGET ngtcp2::ngtcp2_static APPEND PROPERTY IMPORTED_CONFIGURATIONS RELEASE)
set_target_properties(ngtcp2::ngtcp2_static PROPERTIES
  IMPORTED_LINK_INTERFACE_LANGUAGES_RELEASE "C"
  IMPORTED_LOCATION_RELEASE "${_IMPORT_PREFIX}/lib/ngtcp2.lib"
  )

list(APPEND _cmake_import_check_targets ngtcp2::ngtcp2_static )
list(APPEND _cmake_import_check_files_for_ngtcp2::ngtcp2_static "${_IMPORT_PREFIX}/lib/ngtcp2.lib" )

# Commands beyond this point should not need to know the version.
set(CMAKE_IMPORT_FILE_VERSION)
