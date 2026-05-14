from sqlalchemy import create_engine

DATABASE_URL = "sqlite:///ventas.db"

engine = create_engine(DATABASE_URL)