from django.shortcuts import render
from django.views.decorators.csrf import csrf_exempt
from django.utils.crypto import get_random_string
from django.core.mail import send_mail
from django.shortcuts import get_object_or_404
# Create your views here.

def index(request):
    return render(request, 'index.html')

def docpage(request):  
    return render(request, 'doc.html')

def profile(request):  
    return render(request, 'profile.html')

def faq(request):  
    return render(request, 'faq.html')

def login(request):
    return render(request, 'login.html')

def templates(request):
    return render(request, 'templates.html')


def bibliography_template(request):
    return render(request, 'bibliography.html')

def gost_template(request):
    return render(request, 'gost.html')

def diploma_template(request):
    return render(request, 'diploma.html')

def business_template(request):
    return render(request, 'business.html')

def contact_template(request):
    return render(request, 'contact.html')


from .models import *

@csrf_exempt
def send_verification_code(request):
    """
    Author: Vlad Golub
    """
    email = request.GET['email']
   
    if not email:
        return JsonResponse({'status': 'error', 'message': 'Email является обязательным'})
    if '@' not in email:
        return JsonResponse({'status': 'error', 'message': 'Email не является валидным'})
    code = get_random_string(length=6, allowed_chars='0123456789')
    user, created = UserToken.objects.get_or_create(email=email)
    user.code = code
    user.token = None
    user.save()
    html_message = f"""<!DOCTYPE html>
<html style="padding-top: 1rem; padding-bottom: 1rem;">
<body style="font-family: Arial, sans-serif; background-color: #faf5ff; margin: 0; padding: 0;">
    <div style="max-width: 600px; margin: 20px auto; background: linear-gradient(135deg, #ec4899 0%, #8b5cf6 100%); border-radius: 16px; padding: 30px; color: white; text-align: center; box-shadow: 0 8px 25px rgba(139, 92, 246, 0.2);">
        
        <div style="font-size: 28px; font-weight: bold; color: #ffffff; margin-bottom: 20px;">Lexify: Ваш код подтверждения</div>
        <div style="font-size: 16px; line-height: 1.5; color: #fdf4ff;">
            Привет! Спасибо, что выбрали Lexify. Ваш код подтверждения готов:
        </div>
        <div style="background-color: rgba(255, 255, 255, 0.2); padding: 15px; border-radius: 10px; font-size: 32px; letter-spacing: 15px; font-weight: bold; color: #fff; margin: 20px 0; border: 1px solid rgba(255,255,255,0.3);">
            {str(code)}
        </div>

        <div style="font-size: 16px; line-height: 1.5; color: #fdf4ff;">
            Введите этот код на сайте, чтобы продолжить. Если вы не запрашивали код, просто проигнорируйте это письмо.
        </div>
        <div style="margin-top: 20px; font-size: 12px; color: #fae8ff;">
            С любовью, команда Lexify
        </div>
    </div>
</body>
</html>"""

    send_mail(
        f'''{str(code)} код подтверждения от Lexify''',
        f'Ваш код подтверждения: {code}',
        settings.DEFAULT_FROM_EMAIL,
        [email],
        html_message=html_message,
    )

    mail_provider = email.split('@')[-1]
    mail_links = {
        'gmail.com': 'https://mail.google.com',
        'yandex.ru': 'https://mail.yandex.ru',
        'vk.com': 'https://vk.mail.ru',
        'rambler.ru': 'https://mail.rambler.ru',
        'yahoo.com': 'https://mail.yahoo.com',
        'mail.ru': 'https://e.mail.ru',
        'outlook.com': 'https://outlook.live.com/mail'
    }
    mail_link = mail_links.get(mail_provider, '/')

    return JsonResponse({'status': 'ok', 'message': 'Отправили Вам код подтверждения на почту', 'link': mail_link})

@csrf_exempt
def verify_code(request):
    """
    Author: Vlad Golub
    """
    email = request.GET['email']
    code = request.GET['code']

    if not (email and code):
       
        return JsonResponse({'status': 'error', 'message': 'Email и код являются обязательными'})

    user = get_object_or_404(UserToken, email=email)

    if user.code == code:
        token = uuid.uuid4().hex
        user.token = token
        user.code = None
        user.save()

        return JsonResponse({'status': 'ok', 'token': token})

    return JsonResponse({'status': 'error', 'message': 'Неверный код'})


from rest_framework.decorators import api_view
from rest_framework.response import Response
from django.core.files.base import ContentFile
from .models import User, Document
from .auth import authenticate_token
import shutil
import os
import uuid
from rest_framework import status

# ---------- ADMIN ----------


from rest_framework.decorators import api_view
from rest_framework.response import Response
from rest_framework import status
from api.models import User
import uuid

@api_view(["POST"])
def admin_create_user(request):
    admin = authenticate_token(request)
    if not admin.is_admin:
        return Response({"error": "admin only"}, status=403)

    email = request.data.get("email")
    if not email:
        return Response({"error": "email required"}, status=400)

    user, created = User.objects.get_or_create(
        email=email,
        defaults={"token": uuid.uuid4().hex}
    )

    if not created:
        return Response(
            {"error": "user already exists"},
            status=status.HTTP_409_CONFLICT
        )

    return Response(
        {"email": user.email, "user_token": user.token},
        status=status.HTTP_201_CREATED
    )

@csrf_exempt
@api_view(["POST"])
def admin_rotate_user_token(request):
    admin = authenticate_token(request)
    if not admin.is_admin:
        return Response({"error": "admin only"}, status=403)

    email = request.data.get("email")
    if not email:
        return Response({"error": "email required"}, status=400)

    try:
        user = User.objects.get(email=email)
    except User.DoesNotExist:
        return Response({"error": "user not found"}, status=404)

    user.token = uuid.uuid4().hex
    user.save()

    return Response(
        {"email": user.email, "new_token": user.token},
        status=200
    )

@csrf_exempt
@api_view(["GET"])
def admin_get_user(request):
    admin = authenticate_token(request)
    if not admin.is_admin:
        return Response({"error": "admin only"}, status=403)

    email = request.query_params.get("email")
    if not email:
        return Response({"error": "email query parameter required"}, status=400)

    try:
        user = User.objects.get(email=email)
    except User.DoesNotExist:
        return Response({"error": "user not found"}, status=404)

    return Response({
        "email": user.email,
        "token": user.token,
        "is_admin": user.is_admin
    })




from rest_framework.decorators import api_view
from rest_framework.response import Response
from django.utils.timezone import now
from datetime import timedelta

@csrf_exempt
@api_view(["GET"]) 
def admin_get_user_settings(request):
    """
    Admin-only view to get most recent record of style, format, dictionary, skip_pages
    used by a specific user, or default values if none exist.
    """
    # Authenticate and check admin privileges
    admin = authenticate_token(request)
    if not admin.is_admin:
        return Response({"error": "admin only"}, status=403)
    
    # Get email from query parameters
    email = request.query_params.get("email")
    if not email:
        return Response({"error": "email query parameter required"}, status=400)
    
    try:
        user = User.objects.get(email=email)
    except User.DoesNotExist:
        return Response({"error": "user not found"}, status=404)
    
    # Default values
    default_values = {
        "style": "ВШЭ",
        "format": "ГОСТ",
        "dictionary": "Базовый",
        "skip_pages": "0"
    }
    
    # Try to get the most recent document for this user
    try:
        # Get the most recent document (by created_at) for this user
        latest_doc = Document.objects.filter(user=user).order_by('-created_at').first()
        
        if latest_doc:
            # Use values from the most recent document
            result = {
                "email": user.email,
                "style": latest_doc.style if latest_doc.style else default_values["style"],
                "format": latest_doc.format if latest_doc.format else default_values["format"],
                "dictionary": latest_doc.dictionary if latest_doc.dictionary else default_values["dictionary"],
                "skip_pages": latest_doc.skip_pages if latest_doc.skip_pages else default_values["skip_pages"],
                "created_at": latest_doc.created_at,
            }
        else:
            # No documents found, return default values
            result = {
                "email": user.email,
                "style": default_values["style"],
                "format": default_values["format"],
                "dictionary": default_values["dictionary"],
                "skip_pages": default_values["skip_pages"],
                "created_at": None,
            }
        
        return Response(result, status=200)
        
    except Exception as e:
        return Response({"error": str(e)}, status=500)


# ---------- FILE UPLOAD ----------

import os
import gzip
from django.conf import settings
from rest_framework.decorators import api_view
from rest_framework.response import Response
from api.models import Document

@csrf_exempt
@api_view(['POST'])
def upload_file(request):
    user = authenticate_token(request)

    uploaded_file = request.FILES['file']
    original_filename = uploaded_file.name 
    ext = os.path.splitext(uploaded_file.name)[-1]

    skip_pages=request.data.get('skip_pages', '')
    print(skip_pages)

    doc = Document.objects.create(
        user=user,
        style=request.data.get('style', ''),
        format=request.data.get('format', ''),
        dictionary=request.data.get('dictionary', ''),
        skip_pages=request.data.get('skip_pages', ''),
        name=original_filename
    )

    filename = f"{doc.id}{ext}.gz"
    save_dir = os.path.join(settings.BASE_DIR, 'static', 'docs')
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, filename)

    # 3. Сохраняем файл с сжатием gzip
    with gzip.open(save_path, 'wb') as f_out:
        for chunk in uploaded_file.chunks():
            f_out.write(chunk)

    # 4. Обновляем запись в базе с путем до файла и сохраняем
    doc.save()

    return Response({
        "file_token": str(doc.id),
        "status": "uploaded",
        "file_name": filename
    })


@csrf_exempt
@api_view(['GET'])
def file_status(request, file_id):
    user = authenticate_token(request)
    doc = Document.objects.get(id=file_id)

    if doc.user != user:
        return Response({"error": "forbidden"}, status=403)

    return Response({
        "processed": doc.processed,
        "datetime": doc.created_at
    })


import random 

@csrf_exempt
@api_view(['GET'])
def get_work(request):
    admin = authenticate_token(request)
    if not admin.is_admin:
        return Response({"error": "admin only"}, status=403)

    unprocessed_docs = Document.objects.filter(processed=False)

    if unprocessed_docs.exists():
        random_doc = random.choice(list(unprocessed_docs))
        
        return Response({
            "id": str(random_doc.id),
            "style": random_doc.style if hasattr(random_doc, 'style') else "default",
            "format": random_doc.format if hasattr(random_doc, 'format') else "standard",
            "dictionary": random_doc.dictionary if hasattr(random_doc, 'dictionary') else {},
            "skip_pages": random_doc.skip_pages if hasattr(random_doc, 'skip_pages') else [],
            "filename": random_doc.filename if hasattr(random_doc, 'filename') else "",
            "file_url": random_doc.file.url if hasattr(random_doc, 'file') and random_doc.file else None,
            "user": str(random_doc.user.email) if hasattr(random_doc, 'user') and random_doc.user else None,
        })
    else:
        return Response({
            "id": None,  # No docs available
            "message": "No unprocessed documents"
        })
    

@csrf_exempt
@api_view(['POST'])
def admin_change_file_status(request):
    """Admin-only endpoint to change document processed status"""
    
    # Authenticate and check if admin
    admin = authenticate_token(request)
    if not admin.is_admin:
        return Response({"error": "admin only"}, status=403)
    
    # Get parameters from request
    file_id = request.data.get('file_id')
    new_status = request.data.get('processed')
    integer_string = request.data.get('integer_string', "0#0#0#0#0#0#0#0#0#0#0")
    integr_list = integer_string.split("#")
    plagiate = request.data.get('plagiate', " ")
    
    # Validate required parameters
    if not file_id:
        return Response({"error": "file_id is required"}, status=400)
    
    if new_status is None:
        return Response({"error": "processed status is required (true/false)"}, status=400)
    
    try:
        # Convert string to boolean if needed
        if isinstance(new_status, str):
            new_status = new_status.lower() == 'true'
        
        # Get the document
        doc = Document.objects.get(id=file_id)
        
        # Update the status
        doc.processed = bool(new_status)

        doc.count_words = int(integr_list[0])
        doc.count_chars = int(integr_list[1])
        doc.count_sentences = int(integr_list[2])
        doc.count_bad_words = int(integr_list[3])
        doc.count_bad_chars = int(integr_list[4])
        doc.count_bibliography = int(integr_list[5])
        doc.count_bad_bibliography = int(integr_list[6])
        doc.count_not_doi = int(integr_list[7])
        doc.count_suggest_doi = int(integr_list[8])
        doc.count_not_right_bibliography = int(integr_list[9])
        doc.count_styles_error = int(integr_list[10])
        doc.plag = str(plagiate)

        doc.save()
        
        return Response({
            "success": True,
            "message": "File status updated",
            "document_id": str(doc.id),
            "previous_status": not new_status,
            "new_status": doc.processed,
            "user": doc.user.email,
            "updated_at": doc.updated_at
        })
        
    except Document.DoesNotExist:
        return Response({"error": f"Document with ID {file_id} not found"}, status=404)
    except Exception as e:
        return Response({"error": str(e)}, status=500)





from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.views.decorators.http import require_GET
import json

@csrf_exempt
@require_GET
def check_page(request):
    token = request.GET.get("token", "")
    return render(request, "check.html", {"token": token})



from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from .models import UserToken, Document
import uuid

@csrf_exempt
def get_user_documents(request):
    """
    Get all documents for the authenticated user using code from UserToken
    """
    token = request.GET.get('code')
    
    if not token:
        return JsonResponse({'error': 'no code provided'}, status=401)
    
    try:
        user_token = UserToken.objects.get(token=token)
        email = user_token.email
        
        documents = Document.objects.filter(user__email=email).order_by('-created_at')
        
        docs_data = []
        for doc in documents:
            docs_data.append({

                'id': str(doc.id),
                'style': doc.style,
                'format': doc.format,
                'dictionary': doc.dictionary,
                'skip_pages': doc.skip_pages,
                'processed': doc.processed,
                'created_at': doc.created_at,
                'updated_at': doc.updated_at,
                'name': doc.name,
                'count_chars': doc.count_chars,
                'count_words': doc.count_words,
                'count_sentences': doc.count_sentences,
                'count_bad_words': doc.count_bad_words,
                'count_bad_chars': doc.count_bad_chars,
                'count_bibliography': doc.count_bibliography,
                'count_bad_bibliography': doc.count_bad_bibliography,
                'count_not_doi': doc.count_not_doi,
                'count_suggest_doi': doc.count_suggest_doi,
                'count_not_right_bibliography': doc.count_not_right_bibliography,
                'count_styles_error': doc.count_styles_error,

            })
        
        return JsonResponse({
            'status': 'ok',
            'total_documents': documents.count(),
            'documents': docs_data
        })
        
    except UserToken.DoesNotExist:
        return JsonResponse({'error': 'invalid code'}, status=401)
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)

def get_file1(request):
    zip_file = open('/home/sasha2122/applexify2.zip', 'rb')
    response = HttpResponse(zip_file, content_type='application/force-download; charset=utf-8')
    response['Content-Disposition'] = 'attachment; filename="%s"' % 'Applexify.zip'
    return response

def get_file2(request):
    zip_file = open('/home/sasha2122/lexifapi2.zip', 'rb')
    response = HttpResponse(zip_file, content_type='application/force-download; charset=utf-8')
    response['Content-Disposition'] = 'attachment; filename="%s"' % 'Lexifapi.zip'
    return response


# views.py
from django.shortcuts import redirect
from django.contrib.staticfiles import finders
from django.http import HttpResponse, Http404
import os

def favicon_view(request):
    # Find the static file
    favicon_path = finders.find('img/fav.ico')
    
    if favicon_path:
        with open(favicon_path, 'rb') as f:
            icon_data = f.read()
        return HttpResponse(icon_data, content_type='image/png')
    else:
        raise Http404("Favicon not found")

def favicon_view_png(request):
    # Find the static file
    favicon_path = finders.find('img/fav.png')
    
    if favicon_path:
        with open(favicon_path, 'rb') as f:
            icon_data = f.read()
        return HttpResponse(icon_data, content_type='image/png')
    else:
        raise Http404("Favicon not found")

def favicon_view_svg(request):
    # Find the static file
    favicon_path = finders.find('img/fav.svg')
    
    if favicon_path:
        with open(favicon_path, 'rb') as f:
            icon_data = f.read()
        return HttpResponse(icon_data, content_type='image/svg+xml')
    else:
        raise Http404("Favicon not found")





from django.views.decorators.csrf import csrf_exempt
from rest_framework.decorators import api_view
from rest_framework.response import Response
from .models import Document  # Replace with your actual app name

@csrf_exempt
@api_view(['GET'])
def get_all_document_ids(request):
    """Get all document IDs - Admin only"""
    
    # Authenticate and check if admin
    admin = authenticate_token(request)
    if not admin or not admin.is_admin:
        return Response({"error": "admin only"}, status=403)
    
    # Get all document IDs
    document_ids = list(Document.objects.values_list('id', flat=True))
    
    return Response({
        "document_ids": document_ids
    })




from django.shortcuts import render
from django.db.models import Count, Sum, Q, Avg
from django.db.models.functions import TruncDate, TruncHour
from datetime import datetime, timedelta
from .models import Document
import json

def statistics_view(request):
    # Get date range for available data
    first_document = Document.objects.order_by('created_at').first()
    last_document = Document.objects.order_by('created_at').last()
    
    date_range = {
        'first_date': first_document.created_at.date() if first_document else None,
        'last_date': last_document.created_at.date() if last_document else None,
        'total_days': (last_document.created_at.date() - first_document.created_at.date()).days + 1 
                      if first_document and last_document else 0
    }
    
    # 1. STATISTICS BY DAY OF WEEK
    day_of_week_stats = {}
    days_map = {
        1: 'Вс', 2: 'Пн', 3: 'Вт', 4: 'Ср',
        5: 'Чт', 6: 'Пт', 7: 'Сб'
    }
    
    for day_num, day_name in days_map.items():
        day_docs = Document.objects.filter(created_at__week_day=day_num)
        total_docs = day_docs.count()
        total_days_with_data = day_docs.dates('created_at', 'day').distinct().count()
        
        if total_days_with_data > 0:
            day_of_week_stats[day_name] = {
                'total_documents': total_docs,
                'avg_documents_per_day': round(total_docs / total_days_with_data, 2),
                'total_days': total_days_with_data,
                'avg_bad_words': round(day_docs.aggregate(Avg('count_bad_words'))['count_bad_words__avg'] or 0, 2),
                'avg_bad_bibliography': round(day_docs.aggregate(Avg('count_bad_bibliography'))['count_bad_bibliography__avg'] or 0, 2),
                'avg_styles_error': round(day_docs.aggregate(Avg('count_styles_error'))['count_styles_error__avg'] or 0, 2),
            }
    
    # 2. STATISTICS BY HOUR OF DAY
    hour_of_day_stats = {}
    for hour in range(24):
        hour_docs = Document.objects.filter(created_at__hour=hour)
        total_docs = hour_docs.count()
        total_days_with_data = hour_docs.dates('created_at', 'day').distinct().count()
        
        if total_days_with_data > 0:
            hour_of_day_stats[hour] = {
                'total_documents': total_docs,
                'avg_documents_per_day': round(total_docs / total_days_with_data, 2),
                'total_days': total_days_with_data,
                'avg_bad_words': round(hour_docs.aggregate(Avg('count_bad_words'))['count_bad_words__avg'] or 0, 2),
                'avg_bad_bibliography': round(hour_docs.aggregate(Avg('count_bad_bibliography'))['count_bad_bibliography__avg'] or 0, 2),
                'avg_styles_error': round(hour_docs.aggregate(Avg('count_styles_error'))['count_styles_error__avg'] or 0, 2),
            }
    
    # 3. DAILY STATISTICS (for each date with data)
    daily_stats_detail = []
    daily_data = (
        Document.objects
        .annotate(date=TruncDate('created_at'))
        .values('date')
        .annotate(
            total_documents=Count('id'),
            total_bad_words=Sum('count_bad_words'),
            total_bad_bibliography=Sum('count_bad_bibliography'),
            total_styles_error=Sum('count_styles_error'),
            avg_bad_words=Avg('count_bad_words'),
            avg_bad_bibliography=Avg('count_bad_bibliography'),
            avg_styles_error=Avg('count_styles_error'),
        )
        .order_by('-date')
    )
    
    for day in daily_data:
        daily_stats_detail.append({
            'date': day['date'].strftime('%d.%m.%Y') if day['date'] else 'Unknown',
            'documents': day['total_documents'],
            'bad_words': day['total_bad_words'] or 0,
            'bad_bibliography': day['total_bad_bibliography'] or 0,
            'styles_error': day['total_styles_error'] or 0,
            'avg_bad_words': round(day['avg_bad_words'] or 0, 2),
            'avg_bad_bibliography': round(day['avg_bad_bibliography'] or 0, 2),
            'avg_styles_error': round(day['avg_styles_error'] or 0, 2),
        })
    
    # 4. HOURLY BREAKDOWN
    hourly_stats = []
    for hour in range(24):
        hour_docs = Document.objects.filter(created_at__hour=hour)
        stats = hour_docs.aggregate(
            avg_bad_words=Avg('count_bad_words'),
            avg_bad_bibliography=Avg('count_bad_bibliography'),
            avg_styles_error=Avg('count_styles_error'),
            total_docs=Count('id')
        )
        
        hourly_stats.append({
            'hour': f"{hour:02d}:00",
            'documents': stats['total_docs'],
            'avg_bad_words': round(stats['avg_bad_words'] or 0, 2),
            'avg_bad_bibliography': round(stats['avg_bad_bibliography'] or 0, 2),
            'avg_styles_error': round(stats['avg_styles_error'] or 0, 2),
        })
    
    # 5. SUMMARY STATISTICS
    summary_stats = {
        'total_documents_all_time': Document.objects.count(),
        'total_bad_words_all_time': Document.objects.aggregate(Sum('count_bad_words'))['count_bad_words__sum'] or 0,
        'total_bad_bibliography_all_time': Document.objects.aggregate(Sum('count_bad_bibliography'))['count_bad_bibliography__sum'] or 0,
        'total_styles_error_all_time': Document.objects.aggregate(Sum('count_styles_error'))['count_styles_error__sum'] or 0,
        'avg_documents_per_day': round(Document.objects.count() / date_range['total_days'], 2) if date_range['total_days'] > 0 else 0,
        'avg_bad_words_per_document': round(Document.objects.aggregate(Avg('count_bad_words'))['count_bad_words__avg'] or 0, 2),
        'avg_bad_bibliography_per_document': round(Document.objects.aggregate(Avg('count_bad_bibliography'))['count_bad_bibliography__avg'] or 0, 2),
        'avg_styles_error_per_document': round(Document.objects.aggregate(Avg('count_styles_error'))['count_styles_error__avg'] or 0, 2),
    }
    
    # Prepare data for charts
    context = {
        'date_range': date_range,
        'summary_stats': summary_stats,
        'day_of_week_stats': day_of_week_stats,
        'hour_of_day_stats': hour_of_day_stats,
        'daily_stats_detail': daily_stats_detail[:30],  # Last 30 days
        'hourly_stats': hourly_stats,
        
        # JSON data for charts
        'day_of_week_labels': json.dumps(list(day_of_week_stats.keys())),
        'day_of_week_data': json.dumps([stats['avg_documents_per_day'] for stats in day_of_week_stats.values()]),
        'day_of_week_bad_words': json.dumps([stats['avg_bad_words'] for stats in day_of_week_stats.values()]),
        'day_of_week_bad_bib': json.dumps([stats['avg_bad_bibliography'] for stats in day_of_week_stats.values()]),
        'day_of_week_styles': json.dumps([stats['avg_styles_error'] for stats in day_of_week_stats.values()]),
        
        'hour_labels': json.dumps([f"{h:02d}:00" for h in range(24)]),
        'hour_document_data': json.dumps([stats.get('avg_documents_per_day', 0) for stats in hour_of_day_stats.values()]),
        'hour_bad_words_data': json.dumps([stats.get('avg_bad_words', 0) for stats in hour_of_day_stats.values()]),
        'hour_bad_bib_data': json.dumps([stats.get('avg_bad_bibliography', 0) for stats in hour_of_day_stats.values()]),
        'hour_styles_data': json.dumps([stats.get('avg_styles_error', 0) for stats in hour_of_day_stats.values()]),
    }
    
    return render(request, 'stats.html', context)