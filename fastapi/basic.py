from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def index():
    return{"name": "First Date"}

# uvicorn basic:app --reload
# /docs