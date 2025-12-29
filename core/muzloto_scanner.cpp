#include <algorithm>
#include <chrono>
#include <cstring>
#include <iostream>
#include <leptonica/allheaders.h>
#include <map>
#include <memory>
#include <nlohmann/json.hpp>
#include <opencv2/opencv.hpp>
#include <regex>
#include <sstream>
#include <stdlib.h>
#include <string>
#include <tesseract/baseapi.h>
#include <unordered_map>
#include <vector>

using json = nlohmann::json;

#ifdef _WIN32
#define MUZLOTO_EXPORT __declspec(dllexport)
#else
#define MUZLOTO_EXPORT __attribute__((visibility("default")))
#endif

namespace muzloto {

struct FieldResult {
  std::string name;
  std::string value;
  float confidence;
};

struct ScanResult {
  bool success;
  std::string error_message;
  std::vector<FieldResult> fields;
  std::string raw_text;
  double processing_time_ms;

  // 16 полей анкеты
  std::string date;                // 1
  std::string table_number;        // 2
  std::string location;            // 3
  std::string satisfaction_rating; // 4 (1-10)
  std::string playlist_rating;     // 5 (1-10)
  std::string tracks_to_add;       // 6
  std::string location_rating;     // 7 (1-10)
  std::string kitchen_rating;      // 8 (1-10)
  std::string service_rating;      // 9 (1-10)
  std::string host_rating;         // 10 (1-10)
  std::string visits_count;        // 11
  std::string ticket_price;        // 12 (варианты)
  std::string know_booking;        // 13 (Да/Нет)
  std::string source_info;         // 14
  std::string purpose;             // 15
  std::string improvements;        // 16
  std::string phone_number;        // телефон
};

class MUZLOTO_EXPORT MuzlotoScanner {
private:
  std::unique_ptr<tesseract::TessBaseAPI> ocr;
  bool initialized;

  // Точные названия полей из анкеты
  const std::vector<std::pair<std::string, std::string>> field_mapping = {
      {"Дата:", "date"},
      {"Номер столика:", "table_number"},
      {"Место игры:", "location"},
      {"Довольны ли вы посещением Музлото?", "satisfaction_rating"},
      {"Понравился ли вам плейлист?", "playlist_rating"},
      {"Какие треки вы бы добавили?", "tracks_to_add"},
      {"Понравилась ли вам локация?", "location_rating"},
      {"Понравилась ли вам кухня и бар?", "kitchen_rating"},
      {"Устроил ли вас сервис, время подачи?", "service_rating"},
      {"Понравилась ли вам работа ведущего?", "host_rating"},
      {"Сколько раз вы были на Музлото?", "visits_count"},
      {"Оцените стоимость игры за билет", "ticket_price"},
      {"Знаете ли вы, что Музлото можно заказать на корпоратив или день "
       "рождения",
       "know_booking"},
      {"Откуда вы о нас узнали?", "source_info"},
      {"Ради чего вы обычно ходите на подобные вечеринки?", "purpose"},
      {"Что нам стоит улучшить?", "improvements"},
      {"Если вы хотите, чтобы мы с вами связались - оставьте ваш номер "
       "телефона.",
       "phone_number"}};

public:
  MuzlotoScanner() : initialized(false) {
    ocr = std::make_unique<tesseract::TessBaseAPI>();
  }

  ~MuzlotoScanner() {
    if (ocr) {
      ocr->End();
    }
  }

  bool initialize(const std::string &tessdata_path = "") {
    try {
      // Инициализация Tesseract с русским языком
      if (ocr->Init(tessdata_path.empty() ? NULL : tessdata_path.c_str(),
                    "rus+eng", tesseract::OEM_LSTM_ONLY) != 0) {
        return false;
      }

      // Настройки для анкет
      ocr->SetPageSegMode(tesseract::PSM_AUTO);
      ocr->SetVariable("preserve_interword_spaces", "1");
      ocr->SetVariable("textord_tabfind_find_tables", "1");
      ocr->SetVariable("textord_tablefind_recognize_tables", "1");

      initialized = true;
      return true;

    } catch (const std::exception &e) {
      std::cerr << "Ошибка инициализации: " << e.what() << std::endl;
      return false;
    }
  }

  ScanResult scan_image(const std::string &image_path) {
    ScanResult result;
    auto start_time = std::chrono::high_resolution_clock::now();

    try {
      if (!initialized) {
        throw std::runtime_error("Сканер не инициализирован");
      }

      // 1. Загрузка изображения
      cv::Mat image = cv::imread(image_path, cv::IMREAD_COLOR);
      if (image.empty()) {
        throw std::runtime_error("Не удалось загрузить изображение: " +
                                 image_path);
      }

      // 2. Предобработка
      cv::Mat processed = preprocess_image(image);

      // 3. Распознавание текста
      ocr->SetImage(processed.data, processed.cols, processed.rows,
                    processed.channels(), processed.step);

      char *text = ocr->GetUTF8Text();
      result.raw_text = text ? std::string(text) : "";
      delete[] text;

      // 4. Парсинг анкеты Muzloto
      parse_muzloto_form(result);

      // 5. Обработка ответов
      extract_answers(result);

      result.success = true;

    } catch (const std::exception &e) {
      result.success = false;
      result.error_message = e.what();
    }

    auto end_time = std::chrono::high_resolution_clock::now();
    result.processing_time_ms =
        std::chrono::duration<double, std::milli>(end_time - start_time)
            .count();

    return result;
  }

private:
  cv::Mat preprocess_image(const cv::Mat &image) {
    cv::Mat gray, denoised, binary;

    // Конвертация в оттенки серого
    cv::cvtColor(image, gray, cv::COLOR_BGR2GRAY);

    // Удаление шума
    cv::fastNlMeansDenoising(gray, denoised, 10, 7, 21);

    // Улучшение контраста
    cv::Mat equalized;
    cv::equalizeHist(denoised, equalized);

    // Адаптивная бинаризация
    cv::adaptiveThreshold(equalized, binary, 255,
                          cv::ADAPTIVE_THRESH_GAUSSIAN_C, cv::THRESH_BINARY, 11,
                          2);

    return binary;
  }

  void parse_muzloto_form(ScanResult &result) {
    // Разбиваем текст на строки
    std::vector<std::string> lines;
    std::stringstream ss(result.raw_text);
    std::string line;

    while (std::getline(ss, line)) {
      // Очистка строки
      line.erase(std::remove_if(line.begin(), line.end(),
                                [](char c) { return c == '\n' || c == '\r'; }),
                 line.end());

      if (!line.empty()) {
        lines.push_back(line);
      }
    }

    // Словарь для хранения ответов
    std::unordered_map<std::string, std::string> answers;

    // Ищем вопросы и следующие за ними ответы
    for (size_t i = 0; i < lines.size(); i++) {
      std::string current_line = lines[i];

      for (const auto &[question, field_id] : field_mapping) {
        if (current_line.find(question) != std::string::npos) {
          FieldResult field;
          field.name = question;

          // Ищем ответ - следующая непустая строка
          std::string answer_value = "";
          for (size_t j = i + 1; j < lines.size(); j++) {
            if (!lines[j].empty()) {
              // Проверяем, не является ли следующая строка новым вопросом
              bool is_next_question = false;
              for (const auto &[next_q, _] : field_mapping) {
                if (lines[j].find(next_q) != std::string::npos) {
                  is_next_question = true;
                  break;
                }
              }

              if (!is_next_question) {
                answer_value = lines[j];
                i = j; // Пропускаем обработанный ответ
                break;
              }
            }
          }

          field.value = answer_value;
          field.confidence = 0.9f;
          result.fields.push_back(field);
          answers[field_id] = answer_value;
          break;
        }
      }
    }

    // Заполняем структуру
    result.date = answers["date"];
    result.table_number = answers["table_number"];
    result.location = answers["location"];
    result.satisfaction_rating = extract_rating(answers["satisfaction_rating"]);
    result.playlist_rating = extract_rating(answers["playlist_rating"]);
    result.tracks_to_add = answers["tracks_to_add"];
    result.location_rating = extract_rating(answers["location_rating"]);
    result.kitchen_rating = extract_rating(answers["kitchen_rating"]);
    result.service_rating = extract_rating(answers["service_rating"]);
    result.host_rating = extract_rating(answers["host_rating"]);
    result.visits_count = answers["visits_count"];
    result.ticket_price = extract_ticket_price(answers["ticket_price"]);
    result.know_booking = extract_yes_no(answers["know_booking"]);
    result.source_info = answers["source_info"];
    result.purpose = answers["purpose"];
    result.improvements = answers["improvements"];
    result.phone_number = extract_phone_number(answers["phone_number"]);
  }

  void extract_answers(ScanResult &result) {
    // Дополнительная обработка ответов
    if (!result.phone_number.empty()) {
      result.phone_number = normalize_phone(result.phone_number);
    }
  }

  std::string extract_rating(const std::string &text) {
    if (text.empty())
      return "";

    // Ищем числа от 1 до 10
    std::regex rating_regex(R"((10|[1-9]))");
    std::smatch match;

    if (std::regex_search(text, match, rating_regex)) {
      return match.str();
    }

    return text;
  }

  std::string extract_ticket_price(const std::string &text) {
    if (text.empty())
      return "";

    std::string lower_text = text;
    std::transform(lower_text.begin(), lower_text.end(), lower_text.begin(),
                   [](unsigned char c) { return std::tolower(c); });

    if (lower_text.find("можно смело ставить дороже") != std::string::npos ||
        lower_text.find("дороже") != std::string::npos) {
      return "можно смело ставить дороже";
    } else if (lower_text.find("доступно") != std::string::npos) {
      return "доступно";
    } else if (lower_text.find("дорого") != std::string::npos) {
      return "дорого";
    }

    return text;
  }

  std::string extract_yes_no(const std::string &text) {
    if (text.empty())
      return "";

    std::string lower_text = text;
    std::transform(lower_text.begin(), lower_text.end(), lower_text.begin(),
                   [](unsigned char c) { return std::tolower(c); });

    if (lower_text.find("да") != std::string::npos ||
        lower_text.find("yes") != std::string::npos ||
        lower_text.find("✓") != std::string::npos ||
        lower_text.find("+") != std::string::npos ||
        lower_text.find("v") != std::string::npos ||
        lower_text.find("x") != std::string::npos) {
      return "Да";
    } else if (lower_text.find("нет") != std::string::npos ||
               lower_text.find("no") != std::string::npos) {
      return "Нет";
    }

    return text;
  }

  std::string extract_phone_number(const std::string &text) {
    if (text.empty())
      return "";

    // Простая экстракция телефона
    std::regex phone_regex(
        R"((\+7|8)[\s\-\(]?(\d{3})[\s\-\)]?(\d{3})[\s\-]?(\d{2})[\s\-]?(\d{2}))");
    std::smatch match;

    if (std::regex_search(text, match, phone_regex)) {
      return match.str();
    }

    return "";
  }

  std::string normalize_phone(const std::string &phone) {
    std::string normalized = phone;

    // Убираем все нецифровые символы кроме +
    normalized.erase(
        std::remove_if(normalized.begin(), normalized.end(),
                       [](char c) { return !std::isdigit(c) && c != '+'; }),
        normalized.end());

    // Приводим к формату +7XXXXXXXXXX
    if (!normalized.empty()) {
      if (normalized[0] == '8') {
        normalized[0] = '7';
        normalized = "+" + normalized;
      } else if (normalized.substr(0, 2) != "+7") {
        normalized = "+7" + normalized;
      }
    }

    return normalized;
  }
};

} // namespace muzloto

// C-интерфейс для простого использования
extern "C" {
MUZLOTO_EXPORT void *muzloto_create() { return new muzloto::MuzlotoScanner(); }

MUZLOTO_EXPORT void muzloto_destroy(void *scanner) {
  delete static_cast<muzloto::MuzlotoScanner *>(scanner);
}

MUZLOTO_EXPORT int muzloto_initialize(void *scanner,
                                      const char *tessdata_path) {
  return static_cast<muzloto::MuzlotoScanner *>(scanner)->initialize(
             tessdata_path ? std::string(tessdata_path) : "")
             ? 1
             : 0;
}

MUZLOTO_EXPORT const char *muzloto_scan_image(void *scanner,
                                              const char *image_path) {
  try {
    auto result = static_cast<muzloto::MuzlotoScanner *>(scanner)->scan_image(
        image_path ? std::string(image_path) : "");

    // Конвертируем результат в JSON
    nlohmann::json j;
    j["success"] = result.success;
    j["error_message"] = result.error_message;
    j["processing_time_ms"] = result.processing_time_ms;

    // === Поля анкеты Muzloto (16 полей) ===
    j["date"] = result.date;
    j["table_number"] = result.table_number;
    j["location"] = result.location;
    j["satisfaction_rating"] = result.satisfaction_rating;
    j["playlist_rating"] = result.playlist_rating;
    j["tracks_to_add"] = result.tracks_to_add;
    j["location_rating"] = result.location_rating;
    j["kitchen_rating"] = result.kitchen_rating;
    j["service_rating"] = result.service_rating;
    j["host_rating"] = result.host_rating;
    j["visits_count"] = result.visits_count;
    j["ticket_price"] = result.ticket_price;
    j["know_booking"] = result.know_booking;
    j["source_info"] = result.source_info;
    j["purpose"] = result.purpose;
    j["improvements"] = result.improvements;
    j["phone_number"] = result.phone_number;

    j["raw_text"] = result.raw_text.substr(0, 500);

    // Все распознанные поля
    nlohmann::json fields_array = nlohmann::json::array();
    for (const auto &field : result.fields) {
      nlohmann::json f;
      f["name"] = field.name;
      f["value"] = field.value;
      f["confidence"] = field.confidence;
      fields_array.push_back(f);
    }
    j["fields"] = fields_array;

    std::string json_str = j.dump();
    char *c_str = static_cast<char *>(malloc(json_str.length() + 1));
    if (c_str) {
      std::strcpy(c_str, json_str.c_str());
    }
    return c_str;

  } catch (const std::exception &e) {
    nlohmann::json error_json;
    error_json["success"] = false;
    error_json["error_message"] = std::string("C++ exception: ") + e.what();
    error_json["processing_time_ms"] = 0.0;

    std::string error_str = error_json.dump();
    char *error_c_str = static_cast<char *>(malloc(error_str.length() + 1));
    if (error_c_str) {
      std::strcpy(error_c_str, error_str.c_str());
    }
    return error_c_str;
  }
}

MUZLOTO_EXPORT void muzloto_free_string(const char *str) {
  if (str) {
    free(const_cast<char *>(str));
  }
}
}
