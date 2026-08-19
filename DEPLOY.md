# Public deployment

This project is deployment-ready for a Python web host.

Render:
1. Put this project in a GitHub repository.
2. Create a new Web Service on Render.
3. Connect the repository.
4. Build command: `pip install -r requirements.txt`
5. Start command: `gunicorn --bind 0.0.0.0:$PORT app:app`
6. Deploy.
7. Open the generated HTTPS URL on iPhone/iPad/Android/PC.

The included `render.yaml` can also be used with a Blueprint deployment.

Production hardening still recommended:
- HTTPS (provided by the host)
- authentication/login
- rate limiting
- maximum upload size
- automatic deletion of uploaded PDFs/outputs
- persistent storage only if needed
- logging and monitoring
