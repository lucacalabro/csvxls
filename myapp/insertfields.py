from myapp.models import User

for i in ("a", "b", "c", "d"):
    user = User(username="username_" + i, first_name="first_name_" + i, last_name="last_name_" + i, email="email_" + i)
    user.save()
