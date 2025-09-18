import os
import requests
import json
import logging
from pathlib import Path
from fastapi import FastAPI, Request, BackgroundTasks, HTTPException
from dotenv import load_dotenv
from summary import summarize_for_marketing, save_summary
from audiotext import VideoToTextConverter

# Load environment variables from .env
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ----------------- Config -----------------
DOWNLOAD_DIR = "downloads"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

PROCESSED_FILE = "processed.json"  # tracks processed pulses

# Initialize OpenAI converter
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
if not OPENAI_API_KEY:
    raise RuntimeError("❌ OPENAI_API_KEY not found in environment variables")

# Monday.com API configuration
MONDAY_API_KEY = os.getenv("MONDAY_API_KEY")
if MONDAY_API_KEY:
    logger.info("✅ Monday.com API key found")
else:
    logger.warning("⚠️ MONDAY_API_KEY not found - Monday.com updates will be disabled")

converter = VideoToTextConverter(OPENAI_API_KEY)

# ----------------- Monday.com Integration -----------------
def post_update_to_monday(item_id, message, api_key=None):
    """Post an update to Monday.com item"""
    api_key = api_key or MONDAY_API_KEY
    if not api_key:
        logger.warning("⚠️ Monday.com API key not found - skipping update")
        return False
    
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": api_key,
        "Content-Type": "application/json"
    }
    query = """
    mutation ($itemId: Int!, $body: String!) {
      create_update(item_id: $itemId, body: $body) {
        id
      }
    }
    """
    variables = {
        "itemId": item_id,
        "body": message
    }
    
    try:
        response = requests.post(url, json={"query": query, "variables": variables}, headers=headers, timeout=30)
        if response.status_code == 200:
            result = response.json()
            if "errors" in result:
                error_msg = "; ".join([error["message"] for error in result["errors"]])
                logger.error(f"❌ Monday.com API error: {error_msg}")
                return False
            logger.info(f"✅ Update posted to Monday.com item {item_id}")
            return True
        else:
            logger.error(f"❌ Failed to post update: {response.text}")
            return False
    except Exception as e:
        logger.error(f"❌ Error posting to Monday.com: {e}")
        return False

def post_marketing_summary_to_monday(item_id, summary, transcript_length=0):
    """Post formatted marketing summary to Monday.com"""
    timestamp = __import__('datetime').datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    formatted_message = f"""🎯 **MARKETING SUMMARY GENERATED**

📅 **Generated:** {timestamp}
📊 **Source Length:** {transcript_length:,} characters
🤖 **Processed by:** AI Video Analysis System

---

{summary}

---

✅ **Status:** Processing Complete
🎥 **Next Steps:** Review summary and use for marketing content creation"""

    return post_update_to_monday(item_id, formatted_message)

def post_error_to_monday(item_id, error_message, error_type="PROCESSING_ERROR"):
    """Post error message to Monday.com"""
    timestamp = __import__('datetime').datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    formatted_error = f"""❌ **{error_type}**

⏰ **Time:** {timestamp}
🔍 **Error Details:** 
{error_message}

🔧 **Action Required:** Please check the video file and try again, or contact support if the issue persists."""

    return post_update_to_monday(item_id, formatted_error)

# ----------------- Duplicate Tracking -----------------
def is_already_processed(pulse_id):
    if not os.path.exists(PROCESSED_FILE):
        return False
    try:
        with open(PROCESSED_FILE) as f:
            data = json.load(f)
        return pulse_id in data
    except (json.JSONDecodeError, IOError) as e:
        logger.error(f"Error reading processed file: {e}")
        return False

def mark_processed(pulse_id):
    data = {}
    if os.path.exists(PROCESSED_FILE):
        try:
            with open(PROCESSED_FILE) as f:
                data = json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            logger.error(f"Error reading processed file: {e}")
    
    data[pulse_id] = True
    try:
        with open(PROCESSED_FILE, "w") as f:
            json.dump(data, f, indent=2)
        logger.info(f"✅ Marked pulse {pulse_id} as processed")
    except IOError as e:
        logger.error(f"Error writing to processed file: {e}")

# ----------------- Utility Functions -----------------
def convert_google_drive_url(file_url):
    """Convert Google Drive sharing URL to direct download URL"""
    if "drive.google.com/file/d/" in file_url:
        try:
            file_id = file_url.split("/d/")[1].split("/")[0]
            return f"https://drive.google.com/uc?export=download&id={file_id}"
        except IndexError:
            logger.error(f"Could not extract file ID from Google Drive URL: {file_url}")
            return file_url
    return file_url

def download_file(file_url, video_path):
    """Download file with proper error handling and progress logging"""
    try:
        response = requests.get(file_url, stream=True, timeout=300)
        response.raise_for_status()
        
        total_size = int(response.headers.get('content-length', 0))
        downloaded = 0
        
        with open(video_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_size > 0:
                        progress = (downloaded / total_size) * 100
                        if downloaded % (1024 * 1024) == 0:  # Log every MB
                            logger.info(f"📥 Download progress: {progress:.1f}%")
        
        logger.info(f"✅ File downloaded: {video_path} ({downloaded} bytes)")
        return True
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Download failed: {e}")
        return False

# ----------------- Background Task -----------------
def process_video(file_url, pulse_name, pulse_id):
    """Process video file: download -> transcribe -> summarize -> post to Monday.com"""
    try:
        # Skip if already processed
        if is_already_processed(pulse_id):
            logger.info(f"⚠️ Pulse {pulse_id} already processed. Skipping...")
            return

        logger.info(f"🚀 Starting processing for pulse {pulse_id}: {pulse_name}")

        # Convert Google Drive link → direct download
        direct_url = convert_google_drive_url(file_url)
        logger.info(f"📎 Processing URL: {direct_url}")

        # Sanitize filename
        safe_pulse_name = "".join(c for c in pulse_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        video_path = os.path.join(DOWNLOAD_DIR, f"{safe_pulse_name}.mp4")

        # Download video
        logger.info(f"📥 Downloading video to: {video_path}")
        if not download_file(direct_url, video_path):
            error_msg = "Failed to download video file"
            if MONDAY_API_KEY:
                post_error_to_monday(int(pulse_id), error_msg, "DOWNLOAD_ERROR")
            raise Exception(error_msg)

        # Verify file exists and has content
        if not os.path.exists(video_path) or os.path.getsize(video_path) == 0:
            error_msg = "Downloaded file is empty or doesn't exist"
            if MONDAY_API_KEY:
                post_error_to_monday(int(pulse_id), error_msg, "FILE_ERROR")
            raise Exception(error_msg)

        # Transcribe video
        logger.info("🎵 Starting audio extraction and transcription...")
        transcript_file = os.path.join(DOWNLOAD_DIR, f"{safe_pulse_name}_transcript.txt")
        transcript = converter.convert_video_to_text(video_path, transcript_file)
        
        if not transcript or len(transcript.strip()) < 10:
            error_msg = "Transcription failed or produced empty result"
            if MONDAY_API_KEY:
                post_error_to_monday(int(pulse_id), error_msg, "TRANSCRIPTION_ERROR")
            raise Exception(error_msg)
        
        logger.info("✅ Transcription complete!")
        logger.info(f"📄 Transcript preview: {transcript[:200]}...")

        # Generate marketing summary
        logger.info("📝 Generating marketing summary...")
        summary_file = os.path.join(DOWNLOAD_DIR, f"{safe_pulse_name}_marketing_summary.txt")
        summary = summarize_for_marketing(transcript)
        save_summary(summary, summary_file)
        
        logger.info("✅ Marketing summary generated!")
        logger.info(f"📋 Summary preview: {summary[:200]}...")

        # Post summary to Monday.com
        if MONDAY_API_KEY:
            logger.info("📤 Posting marketing summary to Monday.com...")
            success = post_marketing_summary_to_monday(int(pulse_id), summary, len(transcript))
            if success:
                logger.info("✅ Marketing summary posted to Monday.com successfully")
            else:
                logger.warning("⚠️ Failed to post marketing summary to Monday.com")

        # Clean up video file to save space (optional)
        try:
            os.remove(video_path)
            logger.info(f"🗑️ Cleaned up video file: {video_path}")
        except OSError:
            logger.warning(f"⚠️ Could not remove video file: {video_path}")

        # Mark pulse as processed
        mark_processed(pulse_id)
        logger.info(f"🎉 Successfully processed pulse {pulse_id}")

    except Exception as e:
        error_msg = str(e)
        logger.error(f"❌ Error processing pulse {pulse_id}: {error_msg}")
        
        # Post error to Monday.com
        if MONDAY_API_KEY:
            post_error_to_monday(int(pulse_id), error_msg, "PROCESSING_ERROR")

# ----------------- FastAPI App -----------------
app = FastAPI(title="Monday.com Video Processor", version="1.0.0")

@app.get("/")
async def root():
    return {"message": "Monday.com Video Processing Service", "status": "running"}

@app.get("/health")
async def health_check():
    return {
        "status": "healthy",
        "download_dir": DOWNLOAD_DIR,
        "openai_configured": bool(OPENAI_API_KEY),
        "monday_configured": bool(MONDAY_API_KEY)
    }

@app.post("/monday-webhook")
async def monday_webhook_listener(request: Request, background_tasks: BackgroundTasks):
    """Handle Monday.com webhook for video file uploads"""
    try:
        data = await request.json()
        logger.info(f"📨 Received webhook data: {json.dumps(data, indent=2)}")
        
        event = data.get("event", {})
        
        # Extract file information
        new_value = event.get("value", {}) or event.get("previousValue", {})
        file_url = new_value.get("url")
        pulse_name = event.get("pulseName", "downloaded_file").replace(" ", "_")
        pulse_id = event.get("pulseId")
        
        # Validate required fields
        if not file_url:
            logger.warning("⚠️ No URL found in webhook payload")
            return {"status": "error", "message": "No file URL provided"}
        
        if not pulse_id:
            logger.warning("⚠️ No pulse ID found in webhook payload")
            return {"status": "error", "message": "No pulse ID provided"}
        
        logger.info(f"🎯 Processing request - Pulse ID: {pulse_id}, Name: {pulse_name}")
        logger.info(f"🔗 File URL: {file_url}")
        
        # Run video processing in background
        background_tasks.add_task(process_video, file_url, pulse_name, pulse_id)
        
        # Respond immediately to Monday.com
        return {
            "status": "accepted",
            "pulse_id": pulse_id,
            "pulse_name": pulse_name,
            "message": "Video processing started in background"
        }
        
    except json.JSONDecodeError as e:
        logger.error(f"❌ Invalid JSON in webhook: {e}")
        raise HTTPException(status_code=400, detail="Invalid JSON payload")
    
    except Exception as e:
        logger.error(f"❌ Webhook processing error: {e}")
        raise HTTPException(status_code=500, detail=f"Internal server error: {str(e)}")

# Optional: Add endpoint to check processing status
@app.get("/status/{pulse_id}")
async def check_status(pulse_id: str):
    """Check if a pulse has been processed"""
    processed = is_already_processed(pulse_id)
    return {
        "pulse_id": pulse_id,
        "processed": processed,
        "status": "completed" if processed else "pending"
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)