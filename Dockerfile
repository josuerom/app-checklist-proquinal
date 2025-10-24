FROM python:3
WORKDIR /app
COPY . /app
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 9015
RUN python app.py

# ---- IGNORE ----
RUN true <<'COMMENT'
docker run -dit  -p 9015:9014 --name app-checklist-pqn -v "D:\Workspace\Docker:/share" python bash
docker exec -it app-checklist-pqn bash
COMMENT