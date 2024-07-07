# path parameter && query parameter

from fastapi import FastAPI, HTTPException, Path

app = FastAPI()

students = {
    1:{
        "name":"john",
        "age": 17,
        "class":"year 12"
    },
    2:{
        "name":"dones",
        "age": 19,
        "class":"year 19"
    }
}
@app.get('/')
def index():
    return {"name":"First Data"}

@app.get("/get-student/{student_id}")
# specifying datatype of parameter
# if no parameter, no output, catch the error 
def get_student(student_id: int = Path(..., description="The ID of the student you want to view", gt=0, lt=10)):
    if student_id not in students:
        raise HTTPException(status_code=404, detail="Student not found")
    # specific student 
    return students[student_id]

# gt, lt, ge, le 