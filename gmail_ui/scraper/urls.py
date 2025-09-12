from django.urls import path

from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('download/', views.download_excel, name='download_excel'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
]
