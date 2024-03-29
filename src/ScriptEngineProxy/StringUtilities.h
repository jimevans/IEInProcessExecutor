
#ifndef WEBDRIVER_IE_STRINGUTILITIES_H
#define WEBDRIVER_IE_STRINGUTILITIES_H

#include <string>
#include <vector>

class StringUtilities {
private:
  StringUtilities(void);
  ~StringUtilities(void);
public:
  static std::wstring ToWString(const std::string& input);
  static std::string ToString(const std::wstring& input);

  static std::wstring CreateGuid(void);

  static std::string Format(const char* format, ...);
  static std::wstring Format(const wchar_t* format, ...);

  static std::string TrimRight(const std::string& input);
  static std::string TrimLeft(const std::string& input);
  static std::string Trim(const std::string& input);
  static std::wstring TrimRight(const std::wstring& input);
  static std::wstring TrimLeft(const std::wstring& input);
  static std::wstring Trim(const std::wstring& input);

  static void ToBuffer(const std::string& input,
                       std::vector<char>* buffer);
  static void ToBuffer(const std::wstring& input,
                       std::vector<wchar_t>* buffer);

  static void ComposeUnicodeString(std::wstring* input);
  static void DecomposeUnicodeString(std::wstring* input);

  static void Split(const std::string& input,
                    const std::string& delimiter,
                    std::vector<std::string>* tokens);
  static void Split(const std::wstring& input,
                    const std::wstring& delimiter,
                    std::vector<std::wstring>* tokens);

private:
  static void NormalizeUnicodeString(NORM_FORM normalization_form,
                                     std::wstring* input);
};


#endif  // WEBDRIVER_IE_STRINGUTILITIES_H
