from asgiref.wsgi import WsgiToAsgi

from backend.app import app as flask_app

# Render is currently starting the service with `uvicorn main:app`.
# Wrap the Flask app so it can run under an ASGI server.
app = WsgiToAsgi(flask_app)
