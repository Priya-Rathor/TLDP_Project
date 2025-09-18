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
        item_id = data["event"]["pulseId"]   # Item ID comes as int/string
        column_title = data["event"]["columnTitle"]
        new_value = data["event"]["value"]

        # Example: include the link or value in the comment
        comment_text = f"Auto-comment: Column '{column_title}' was updated with value: {new_value}"

        query = """
        mutation ($item_id: ID!, $body: String!) {
          create_update (item_id: $item_id, body: $body) {
            id
          }
        }
        """

        variables = {
            "item_id": str(item_id),   # 👈 Pass as string (ID)
            "body": comment_text
        }

        headers = {
            "Authorization": MONDAY_API_KEY,
            "Content-Type": "application/json"
        }

        response = requests.post(
            MONDAY_API_URL,
            json={"query": query, "variables": variables},
            headers=headers
        )

        print("✅ Comment added:", response.json())

    except Exception as e:
        print("⚠️ Error:", e)

    return {"status": "ok"}