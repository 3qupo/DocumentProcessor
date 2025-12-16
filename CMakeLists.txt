cmake_minimum_required(VERSION 3.15)
project(DocumentProcessor VERSION 1.0.0 LANGUAGES CXX)

# Настройка стандарта C++ (C++17 для кросс-платформенности)
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Политики CMake для корректной работы на всех ОС
cmake_policy(SET CMP0077 NEW)  # Для find_package

# Опции для пользователя
option(BUILD_SHARED_LIBS "Построение shared библиотек" OFF)
option(WITH_GUI "Построение с GUI интерфейсом" ON)
option(WITH_TESTS "Включить тесты" OFF)
option(WITH_PYTHON_BINDINGS "Создать Python биндинги" OFF)

# Платформозависимые настройки
if(WIN32)
    # Windows специфичные настройки
    add_definitions(-DNOMINMAX -D_USE_MATH_DEFINES -DWIN32_LEAN_AND_MEAN)
    
    # Отключаем предупреждения для внешних библиотек
    add_compile_options(/W4 /WX /MP)
    
    # Для поддержки UTF-8 на Windows
    add_compile_definitions(_UNICODE UNICODE)
    
    # Пути для Windows
    set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)
    set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
    set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
    
elseif(APPLE)
    # macOS специфичные настройки
    add_compile_options(-Wall -Wextra -Wpedantic -Werror)
    
    # Для Homebrew пакетов
    if(EXISTS /usr/local/opt)
        list(APPEND CMAKE_PREFIX_PATH /usr/local/opt)
    endif()
    
    # Требования для macOS
    set(CMAKE_MACOSX_RPATH ON)
    set(CMAKE_INSTALL_RPATH "@executable_path/../lib")
    
elseif(UNIX)
    # Linux специфичные настройки
    add_compile_options(-Wall -Wextra -Wpedantic -Werror -fPIC)
    
    # Позиционирование выходных файлов
    set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR})
    set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR})
    set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR})
endif()

# Настройка поиска зависимостей
set(CMAKE_MODULE_PATH ${CMAKE_MODULE_PATH} "${CMAKE_SOURCE_DIR}/cmake")

# Поиск OpenCV (кросс-платформенно)
find_package(OpenCV REQUIRED
    COMPONENTS core imgproc imgcodecs highgui  # базовые компоненты
)

if(NOT OpenCV_FOUND)
    message(WARNING "OpenCV не найден. Попробуйте установить:")
    if(WIN32)
        message("  - vcpkg install opencv[contrib]:x64-windows")
        message("  - Или скачайте с opencv.org")
    elseif(APPLE)
        message("  - brew install opencv")
    elseif(UNIX)
        message("  - sudo apt-get install libopencv-dev  # Ubuntu/Debian")
        message("  - sudo yum install opencv-devel       # Fedora/RHEL")
    endif()
    message(FATAL_ERROR "OpenCV required для проекта")
endif()

message(STATUS "OpenCV найдена: ${OpenCV_VERSION}")
message(STATUS "OpenCV библиотеки: ${OpenCV_LIBS}")

# Поиск Tesseract OCR
find_package(PkgConfig)
if(PKG_CONFIG_FOUND)
    pkg_check_modules(TESSERACT QUIET tesseract IMPORTED_TARGET)
    pkg_check_modules(LEPTONICA QUIET lept IMPORTED_TARGET)
endif()

if(TESSERACT_FOUND AND LEPTONICA_FOUND)
    message(STATUS "Tesseract найдена через pkg-config: ${TESSERACT_VERSION}")
    set(HAVE_TESSERACT TRUE)
else()
    # Альтернативный поиск
    find_path(TESSERACT_INCLUDE_DIR tesseract/baseapi.h
        PATHS
        /usr/include
        /usr/local/include
        /opt/local/include
        /opt/homebrew/include
        "C:/Program Files/Tesseract-OCR/include"
    )
    
    find_library(TESSERACT_LIBRARY NAMES tesseract libtesseract
        PATHS
        /usr/lib
        /usr/local/lib
        /opt/local/lib
        /opt/homebrew/lib
        "C:/Program Files/Tesseract-OCR/lib"
    )
    
    find_library(LEPTONICA_LIBRARY NAMES lept leptonica
        PATHS
        /usr/lib
        /usr/local/lib
        /opt/local/lib
        /opt/homebrew/lib
        "C:/Program Files/Tesseract-OCR/lib"
    )
    
    if(TESSERACT_INCLUDE_DIR AND TESSERACT_LIBRARY AND LEPTONICA_LIBRARY)
        set(HAVE_TESSERACT TRUE)
        message(STATUS "Tesseract найдена: ${TESSERACT_LIBRARY}")
    else()
        set(HAVE_TESSERACT FALSE)
        message(WARNING "Tesseract не найдена. OCR функции будут отключены.")
        message("Установите Tesseract:")
        if(WIN32)
            message("  - Скачайте с github.com/UB-Mannheim/tesseract/wiki")
        elseif(APPLE)
            message("  - brew install tesseract")
        elseif(UNIX)
            message("  - sudo apt-get install libtesseract-dev  # Ubuntu")
            message("  - sudo yum install tesseract-devel       # Fedora")
        endif()
    endif()
endif()

# Поиск дополнительных библиотек
find_package(ZLIB REQUIRED)

# JSON для конфигураций (опционально)
find_package(nlohmann_json QUIET)
if(NOT nlohmann_json_FOUND)
    # Используем встроенный или FetchContent
    message(STATUS "Используем FetchContent для nlohmann_json")
    include(FetchContent)
    FetchContent_Declare(
        nlohmann_json
        GIT_REPOSITORY https://github.com/nlohmann/json.git
        GIT_TAG v3.11.2
    )
    FetchContent_MakeAvailable(nlohmann_json)
endif()

# SQLite для хранения данных (опционально)
find_package(SQLite3 QUIET)
if(NOT SQLite3_FOUND)
    # Встроенный SQLite
    add_subdirectory(thirdparty/sqlite)
    set(SQLite3_LIBRARY sqlite)
endif()

# Конфигурация GUI
if(WITH_GUI)
    if(WIN32)
        find_package(Qt6 COMPONENTS Core Widgets REQUIRED)
    else()
        # Ищем Qt через pkg-config или стандартный путь
        find_package(Qt6 QUIET COMPONENTS Core Widgets)
        if(NOT Qt6_FOUND)
            find_package(Qt5 COMPONENTS Core Widgets REQUIRED)
        endif()
    endif()
endif()

# Добавление исходников
set(SOURCES
    src/main.cpp
    src/document_scanner.cpp
    src/image_processor.cpp
    src/ocr_engine.cpp
    src/data_exporter.cpp
    src/utils.cpp
)

set(HEADERS
    include/document_scanner.h
    include/image_processor.h
    include/ocr_engine.h
    include/data_exporter.h
    include/utils.h
    include/config.h
)

# Создание исполняемого файла
add_executable(${PROJECT_NAME} ${SOURCES} ${HEADERS})

# Настройка целевых свойств
target_include_directories(${PROJECT_NAME} PRIVATE 
    ${CMAKE_CURRENT_SOURCE_DIR}/include
    ${OpenCV_INCLUDE_DIRS}
)

if(HAVE_TESSERACT)
    if(TARGET PkgConfig::TESSERACT)
        target_include_directories(${PROJECT_NAME} PRIVATE 
            ${TESSERACT_INCLUDE_DIRS}
        )
    else()
        target_include_directories(${PROJECT_NAME} PRIVATE 
            ${TESSERACT_INCLUDE_DIR}
        )
    endif()
    
    # Добавляем определение для условной компиляции
    target_compile_definitions(${PROJECT_NAME} PRIVATE HAVE_TESSERACT)
endif()

# Линковка библиотек
target_link_libraries(${PROJECT_NAME} PRIVATE 
    ${OpenCV_LIBS}
    ${ZLIB_LIBRARIES}
)

if(HAVE_TESSERACT)
    if(TARGET PkgConfig::TESSERACT)
        target_link_libraries(${PROJECT_NAME} PRIVATE 
            PkgConfig::TESSERACT 
            PkgConfig::LEPTONICA
        )
    else()
        target_link_libraries(${PROJECT_NAME} PRIVATE 
            ${TESSERACT_LIBRARY}
            ${LEPTONICA_LIBRARY}
        )
    endif()
endif()

if(WITH_GUI)
    if(TARGET Qt6::Core AND TARGET Qt6::Widgets)
        target_link_libraries(${PROJECT_NAME} PRIVATE 
            Qt6::Core 
            Qt6::Widgets
        )
    elseif(TARGET Qt5::Core AND TARGET Qt5::Widgets)
        target_link_libraries(${PROJECT_NAME} PRIVATE 
            Qt5::Core 
            Qt5::Widgets
        )
    endif()
    
    target_compile_definitions(${PROJECT_NAME} PRIVATE WITH_GUI)
endif()

if(SQLite3_FOUND)
    target_link_libraries(${PROJECT_NAME} PRIVATE SQLite::SQLite3)
elseif(TARGET sqlite)
    target_link_libraries(${PROJECT_NAME} PRIVATE sqlite)
endif()

if(TARGET nlohmann_json)
    target_link_libraries(${PROJECT_NAME} PRIVATE nlohmann_json::nlohmann_json)
endif()

# Платформозависимые библиотеки
if(WIN32)
    target_link_libraries(${PROJECT_NAME} PRIVATE 
        ws2_32
        advapi32
        shell32
    )
    
    # Для Tesseract на Windows
    if(HAVE_TESSERACT AND NOT TARGET PkgConfig::TESSERACT)
        target_link_libraries(${PROJECT_NAME} PRIVATE 
            wsock32
        )
    endif()
    
elseif(APPLE)
    # macOS frameworks
    find_library(FOUNDATION Foundation)
    find_library(APPKIT AppKit)
    target_link_libraries(${PROJECT_NAME} PRIVATE 
        ${FOUNDATION}
        ${APPKIT}
    )
    
elseif(UNIX AND NOT APPLE)
    # Linux системные библиотеки
    target_link_libraries(${PROJECT_NAME} PRIVATE 
        pthread
        dl
        m
    )
endif()

# Копирование ресурсов и данных
configure_file(${CMAKE_SOURCE_DIR}/config/config.json.in 
               ${CMAKE_BINARY_DIR}/config.json @ONLY)

# Копирование тестовых изображений
file(GLOB TEST_IMAGES "${CMAKE_SOURCE_DIR}/data/test_images/*")
file(COPY ${TEST_IMAGES} DESTINATION ${CMAKE_BINARY_DIR}/test_images)

# Установка
install(TARGETS ${PROJECT_NAME}
    RUNTIME DESTINATION bin
    BUNDLE DESTINATION .
)

install(DIRECTORY ${CMAKE_SOURCE_DIR}/data/tessdata
    DESTINATION share/${PROJECT_NAME}
    PATTERN "*.traineddata"
)

# Тесты (опционально)
if(WITH_TESTS)
    enable_testing()
    add_subdirectory(tests)
endif()

# Python биндинги (опционально)
if(WITH_PYTHON_BINDINGS)
    find_package(Python3 COMPONENTS Interpreter Development REQUIRED)
    find_package(pybind11 REQUIRED)
    
    pybind11_add_module(pydocumentprocessor MODULE src/python_bindings.cpp)
    target_link_libraries(pydocumentprocessor PRIVATE ${PROJECT_NAME})
    
    install(TARGETS pydocumentprocessor
        DESTINATION ${Python3_SITEARCH}
    )
endif()

# Информация о сборке
message(STATUS "")
message(STATUS "=== Конфигурация DocumentProcessor ===")
message(STATUS "Система: ${CMAKE_SYSTEM_NAME}")
message(STATUS "Процессор: ${CMAKE_SYSTEM_PROCESSOR}")
message(STATUS "Компилятор: ${CMAKE_CXX_COMPILER_ID}")
message(STATUS "OpenCV: ${OpenCV_VERSION}")
message(STATUS "Tesseract: ${HAVE_TESSERACT}")
message(STATUS "GUI: ${WITH_GUI}")
message(STATUS "Выходная директория: ${CMAKE_BINARY_DIR}")
message(STATUS "=====================================")
