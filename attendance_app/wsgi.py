from app import app, init_db

# Ensure database tables exist when the service boots.
init_db()
