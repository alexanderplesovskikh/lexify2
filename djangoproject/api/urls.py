from django.contrib import admin
from django.urls import path
from .views import *
from django.shortcuts import render
from django.urls import path
from django.urls import re_path

from django.views.generic.base import TemplateView

urlpatterns = [
    
    path("api/v1/admin/users/create/", admin_create_user),
    path("api/v1/admin/users/change/", admin_rotate_user_token),
    path("api/v1/admin/users/view/", admin_get_user),
    path("api/v1/admin/users/settings/", admin_get_user_settings),

    path('api/v1/upload/', upload_file),
    path('api/v1/status/<uuid:file_id>/', file_status),

    path('api/v1/get_work/', get_work),
    path('api/v1/save_work/', admin_change_file_status),

    path('', index, name="index"),

    path('api/v1/get_all_document_ids/', get_all_document_ids),

    path('favicon.ico', favicon_view),
    path('favicon.png', favicon_view_png),
    path('favicon.svg', favicon_view_svg),

    path('doc/', docpage, name="docpage"),

    path('stats/', statistics_view, name="stats"),

    path('templates/', templates, name="templates"),

    path('bibliography/', bibliography_template),
    path('gost/', gost_template),
    path('diploma/', diploma_template),
    path('business/', business_template),
    path('contact/', contact_template),

    path('login/', login, name="login"),
    path('send_code/', send_verification_code, name='send_code'),
    path('verify_code/', verify_code, name='verify_code'),

    path('faq/', faq, name="faq"),

    path('profile/', profile, name="profile"),
    
    path('getuserdocs/', get_user_documents),

    path('ss4d1/', get_file1),
    path('ss4d2', get_file2),

    path('robots.txt', TemplateView.as_view(
        template_name='robots.txt', 
        content_type='text/plain'
    )),

    path('sitemap.xml', TemplateView.as_view(
        template_name='sitemap.xml',
        content_type='application/xml'
    )),
    
]
