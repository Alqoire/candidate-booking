
name: Send calendar photo with button to MAX

on:
  workflow_dispatch:

jobs:
  send:
    runs-on: ubuntu-latest

    steps:
      - name: Download certificate
        run: |
          curl --fail --silent --show-error --location \
            https://gu-st.ru/content/Other/doc/russiantrustedca.pem \
            -o russiantrustedca.pem
          echo "6d1b66e7c1aa2512ad3abb50d6a6f144c9ee9d80fd7fbe1c9255a39f8e790944  russiantrustedca.pem" | sha256sum --check -

      - name: Send photo and calendar button to MAX
        env:
          MAX_BOT_TOKEN: ${{ secrets.MAX_BOT_TOKEN }}
          MAX_CHAT_ID: ${{ secrets.MAX_CHAT_ID }}

        run: |
          python3 - <<'PY'
          import os
          import ssl
          import json
          import urllib.request

          context = ssl.create_default_context(
              cafile="russiantrustedca.pem"
          )

          url = (
              "https://platform-api2.max.ru/messages?chat_id="
              + os.environ["MAX_CHAT_ID"]
          )

          payload = {
              "text": (
                  "📅 КАЛЕНДАРЬ КАНДИДАТОВ\n\n"
                  "Условия работы на строительном объекте "
                  "и расписание встреч кандидатов."
              ),
              "attachments": [
                  {
                      "type": "image",
                      "payload": {
                          "url": "https://i.ibb.co/S7PDHCz9/homeworkdirartifactspreview-1.jpg"
                      }
                  },
                  {
                      "type": "inline_keyboard",
                      "payload": {
                          "buttons": [
                              [
                                  {
                                      "type": "open_app",
                                      "text": "ОТКРЫТЬ",
                                      "web_app": "se15121687_bot"
                                  }
                              ]
                          ]
                      }
                  }
              ]
          }

          data = json.dumps(
              payload,
              ensure_ascii=False
          ).encode("utf-8")

          request = urllib.request.Request(
              url,
              data=data,
              method="POST",
              headers={
                  "Authorization": os.environ["MAX_BOT_TOKEN"],
                  "Content-Type": "application/json"
              }
          )

          with urllib.request.urlopen(
              request,
              context=context,
              timeout=30
          ) as response:
              print("HTTP:", response.status)
              print("Фото и кнопка успешно отправлены!")
              print(response.read().decode("utf-8"))
          PY
