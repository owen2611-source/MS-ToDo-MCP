FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY microsoft_todo_mcp_server.py .
EXPOSE 3000
CMD ["python", "microsoft_todo_mcp_server.py"]
