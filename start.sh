cd /home/sasha2122/lexifapi/venv
source bin/activate
cd ..
cd djangoproject
python3 manage.py makemigrations
python3 manage.py migrate
nohup python3 manage.py runserver > site.log 2>&1 &
cd ..
#nohup python3 telegrambot.py > tg.log 2>&1 &
nohup python3 loop.py > loop.log 2>&1 &
#nohup python3 monitor.py > monitor.log 2>&1 &
#nohup python3 newtg.py > newtg.log 2>&1 & 
pip freeze > require.txt
deactivate
cd
cd /home/sasha2122/applexify/venv
source bin/activate
cd ..
cd django_project
python3 manage.py makemigrations
python3 manage.py migrate
nohup python3 manage.py runserver 9000 > app.log 2>&1 &
pip freeze > require.txt
