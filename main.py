from fastapi import FastAPI
import api
app = FastAPI(title= "Employee Attendance")

app.include_router(api.router)
