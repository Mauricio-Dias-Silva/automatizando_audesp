from django.urls import path
from . import views

app_name = 'main'

urlpatterns = [
    path('', views.index, name='index'),
    path('create-csv/', views.create_csv_view, name='create_csv'),
    path('process/', views.process_view, name='process'),
]
