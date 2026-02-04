"""
Debug and test script for greysage-dashboard API
Triggers the handler from api/process.py
"""

import os
import sys
from pathlib import Path
from http.server import HTTPServer

# Add api directory to path
api_dir = Path(__file__).parent / "api"
sys.path.insert(0, str(api_dir))

# Set up environment variables for testing
os.environ["ONEDRIVE_FILE_URL"] = os.getenv("ONEDRIVE_FILE_URL", "")

# Import the process module
try:
    import process
    print("✓ Successfully imported process module")
except ImportError as e:
    print(f"✗ Failed to import process module: {e}")
    sys.exit(1)


def start_server(port=8000):
    """Start HTTP server with the handler"""
    print("\n" + "="*70)
    print(f"STARTING SERVER ON PORT {port}")
    print("="*70)
    print(f"✓ Handler class: {process.handler}")
    print(f"✓ Maker sheets: {process.MAKER_SHEETS}")
    print(f"✓ OneDrive URL set: {'Yes' if os.getenv('ONEDRIVE_FILE_URL') else 'No'}")
    print("\nStarting HTTP server...")
    print(f"Access endpoint at: http://localhost:{port}/api")
    print("\nPress Ctrl+C to stop the server\n")
    
    try:
        server = HTTPServer(('localhost', port), process.handler)
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n✓ Server stopped")
        server.server_close()
    except Exception as e:
        print(f"✗ Server error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
    start_server(port)
