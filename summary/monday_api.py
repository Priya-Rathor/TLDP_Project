import os
import requests
import json
import logging
from typing import Optional, Dict, Any

logger = logging.getLogger(__name__)

class MondayAPI:
    """
    Monday.com API integration for posting updates and managing items.
    """
    
    def __init__(self, api_key: Optional[str] = None):
        """
        Initialize Monday.com API client.
        
        Args:
            api_key: Monday.com API key (optional, will use env var if not provided)
        """
        self.api_key = api_key or os.getenv("MONDAY_API_KEY")
        if not self.api_key:
            raise ValueError("Monday.com API key not found in environment variables or parameters")
        
        self.base_url = "https://api.monday.com/v2"
        self.headers = {
            "Authorization": self.api_key,
            "Content-Type": "application/json"
        }
        logger.info("🔗 Monday.com API client initialized")

    def post_update_to_item(self, item_id: int, message: str) -> bool:
        """
        Post an update/comment to a Monday.com item.
        
        Args:
            item_id: Monday.com item/pulse ID
            message: Update message to post
        
        Returns:
            bool: True if successful, False otherwise
        """
        query = """
        mutation ($itemId: Int!, $body: String!) {
          create_update(item_id: $itemId, body: $body) {
            id
            body
            created_at
          }
        }
        """
        
        variables = {
            "itemId": item_id,
            "body": message
        }
        
        try:
            logger.info(f"📤 Posting update to Monday.com item {item_id}")
            
            response = requests.post(
                self.base_url,
                json={"query": query, "variables": variables},
                headers=self.headers,
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                
                if "errors" in result:
                    error_msg = "; ".join([error["message"] for error in result["errors"]])
                    logger.error(f"❌ Monday.com API error: {error_msg}")
                    return False
                
                update_data = result.get("data", {}).get("create_update", {})
                update_id = update_data.get("id")
                
                logger.info(f"✅ Update posted successfully to Monday.com item {item_id} (Update ID: {update_id})")
                return True
            else:
                logger.error(f"❌ Failed to post update: HTTP {response.status_code} - {response.text}")
                return False
                
        except requests.exceptions.RequestException as e:
            logger.error(f"❌ Network error posting to Monday.com: {e}")
            return False
        except Exception as e:
            logger.error(f"❌ Unexpected error posting to Monday.com: {e}")
            return False

    def post_marketing_summary(self, item_id: int, summary: str, transcript_length: int = 0) -> bool:
        """
        Post a formatted marketing summary as an update to Monday.com.
        
        Args:
            item_id: Monday.com item/pulse ID
            summary: Marketing summary text
            transcript_length: Length of original transcript (for metadata)
        
        Returns:
            bool: True if successful, False otherwise
        """
        # Format the summary for Monday.com
        formatted_message = self._format_summary_for_monday(summary, transcript_length)
        
        return self.post_update_to_item(item_id, formatted_message)

    def _format_summary_for_monday(self, summary: str, transcript_length: int = 0) -> str:
        """
        Format marketing summary for Monday.com update with proper styling.
        
        Args:
            summary: Raw marketing summary
            transcript_length: Character count of original transcript
        
        Returns:
            str: Formatted message for Monday.com
        """
        timestamp = __import__('datetime').datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Clean up the summary and add Monday.com friendly formatting
        formatted_summary = f"""🎯 **MARKETING SUMMARY GENERATED**

📅 **Generated:** {timestamp}
📊 **Source Length:** {transcript_length:,} characters
🤖 **Processed by:** AI Video Analysis System

---

{summary}

---

✅ **Status:** Processing Complete
🎥 **Next Steps:** Review summary and use for marketing content creation"""

        return formatted_summary

    def post_processing_status(self, item_id: int, status: str, details: str = "") -> bool:
        """
        Post processing status updates to Monday.com item.
        
        Args:
            item_id: Monday.com item/pulse ID
            status: Status message (e.g., "STARTED", "COMPLETED", "ERROR")
            details: Additional details about the status
        
        Returns:
            bool: True if successful, False otherwise
        """
        status_icons = {
            "STARTED": "🚀",
            "DOWNLOADING": "📥",
            "EXTRACTING_AUDIO": "🎵",
            "TRANSCRIBING": "📝",
            "SUMMARIZING": "🎯",
            "COMPLETED": "✅",
            "ERROR": "❌",
            "WARNING": "⚠️"
        }
        
        icon = status_icons.get(status.upper(), "ℹ️")
        timestamp = __import__('datetime').datetime.now().strftime("%H:%M:%S")
        
        message = f"{icon} **{status.upper()}** [{timestamp}]"
        if details:
            message += f"\n{details}"
        
        return self.post_update_to_item(item_id, message)

    def get_item_info(self, item_id: int) -> Optional[Dict[str, Any]]:
        """
        Get information about a Monday.com item.
        
        Args:
            item_id: Monday.com item/pulse ID
        
        Returns:
            dict: Item information or None if failed
        """
        query = """
        query ($itemId: [Int!]!) {
          items(ids: $itemId) {
            id
            name
            board {
              id
              name
            }
            column_values {
              id
              text
              value
            }
          }
        }
        """
        
        variables = {"itemId": [item_id]}
        
        try:
            response = requests.post(
                self.base_url,
                json={"query": query, "variables": variables},
                headers=self.headers,
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                items = result.get("data", {}).get("items", [])
                return items[0] if items else None
            else:
                logger.error(f"❌ Failed to get item info: {response.text}")
                return None
                
        except Exception as e:
            logger.error(f"❌ Error getting item info: {e}")
            return None

    def post_error_update(self, item_id: int, error_message: str, error_type: str = "PROCESSING_ERROR") -> bool:
        """
        Post an error update to Monday.com item.
        
        Args:
            item_id: Monday.com item/pulse ID
            error_message: Error message to post
            error_type: Type of error
        
        Returns:
            bool: True if successful, False otherwise
        """
        timestamp = __import__('datetime').datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        formatted_error = f"""❌ **{error_type}**

⏰ **Time:** {timestamp}
🔍 **Error Details:** 
{error_message}

🔧 **Action Required:** Please check the video file and try again, or contact support if the issue persists."""

        return self.post_update_to_item(item_id, formatted_error)

# Utility functions for backward compatibility
def post_update_to_monday(item_id: int, message: str, api_key: str) -> bool:
    """
    Legacy function for backward compatibility.
    
    Args:
        item_id: Monday.com item ID
        message: Update message
        api_key: Monday.com API key
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        monday_api = MondayAPI(api_key)
        return monday_api.post_update_to_item(item_id, message)
    except Exception as e:
        logger.error(f"❌ Error in legacy post_update_to_monday: {e}")
        return False

def post_marketing_summary_to_monday(item_id: int, summary: str, api_key: str, 
                                   transcript_length: int = 0) -> bool:
    """
    Post marketing summary to Monday.com with proper formatting.
    
    Args:
        item_id: Monday.com item ID
        summary: Marketing summary text
        api_key: Monday.com API key
        transcript_length: Original transcript length
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        monday_api = MondayAPI(api_key)
        return monday_api.post_marketing_summary(item_id, summary, transcript_length)
    except Exception as e:
        logger.error(f"❌ Error posting marketing summary to Monday.com: {e}")
        return False