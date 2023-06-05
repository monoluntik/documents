import requests

def translate_text(text, target_language='en'):


    url = "https://translate.yandex.net/api/v1.5/tr.json/translate"
    params = {
        "key": 'trnsl.1.1.20230522T072555Z.9edbd15635035de3.667394495a8817dc802516f43988d41e32527475',
        "text": text,
        "lang": target_language
    }

    response = requests.get(url, params=params)
    if response.status_code == 200:
        translation = response.json()["text"][0]
        return translation
    else:
        return None

# # Пример использования
# text_to_translate = "Привет"
# target_language = "en"  # Целевой язык перевода (в данном случае, русский)

# translation = translate_text(text_to_translate, target_language)
# print(translation)

