import requests
from fastapi import FastAPI, Request

app = FastAPI()

MONDAY_API_KEY = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjU0NjI5MjM1NywiYWFpIjoxMSwidWlkIjo3NDc3Njk5NywiaWFkIjoiMjAyNS0wOC0wNFQwOTo0MzowNS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MTIxNDMyMDQsInJnbiI6InVzZTEifQ.yYeelRXHOZlaxwYHBAvi6eXRzD2fNn1H-jX-Pd8Ukcw"   # 🔑 Replace with your API Key
MONDAY_API_URL = "https://api.monday.com/v2"

@app.post("/monday-webhook")
async def monday_webhook(request: Request):
    data = await request.json()
    print("🔔 Webhook event received:", data)

    try:
        board_id = data["event"]["boardId"]

        query = """
        query ($board_id: [ID!]) {
          boards (ids: $board_id) {
            id
            name
            state
            description
            workspace {
              id
              name
            }
            groups {
              id
              title
            }
            columns {
              id
              title
              type
            }
          }
        }
        """

        variables = {"board_id": [str(board_id)]}  # 👈 Pass as string list

        headers = {
            "Authorization": MONDAY_API_KEY,
            "Content-Type": "application/json"
        }

        response = requests.post(
            MONDAY_API_URL,
            json={"query": query, "variables": variables},
            headers=headers
        )

        board_details = response.json()
        print("📋 Project (Board) Details:", board_details)

    except Exception as e:
        print("⚠️ Error:", e)

    return {"status": "ok"}