from django.urls import path
from . import views

urlpatterns = [
    path('', views.index),
    path('submit_form', views.submit_form, name="submit_form"),
    path('success', views.success, name="success"),
]
