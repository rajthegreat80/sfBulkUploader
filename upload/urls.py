from django.conf.urls import url
from . import views

urlpatterns = [
    url(r'^$', views.index, name = "index"),
    url(r'^save$', views.upload_file, name = "upload_file"),
    url(r'^uploadTemplate.xlsx', views.send_file, name = "send_file"),
]
