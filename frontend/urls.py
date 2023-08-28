from django.urls import path
from . import views

urlpatterns = [
    path('index', views.index),
    path('', views.index2),
    path('submit_form', views.submit_form, name="submit_form"),
    path('success', views.success, name="success"),
    path('download_users_data', views.download_users_data, name="download_users_data"),
]
