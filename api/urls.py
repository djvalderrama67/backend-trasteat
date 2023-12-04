from django.urls import path, include
from rest_framework import routers
from api import views
from .models import *
from .serializer import *
from api.views import apiOverview

router = routers.DefaultRouter()
router.register(r'categoria', views.CategoriaViewSet)
router.register(r'vehiculo', views.VehiculoViewSet)
router.register(r'bodega', views.BodegaViewSet)
router.register(r'objeto', views.ObjetoViewSet)
router.register(r'calculo', views.CalculoViewSet)


urlpatterns = [
    path('api/overview/', apiOverview, name='api-overview'),
    path('categoria/<str:nombre_categoria>/', views.CategoriaViewSet.as_view({'get': 'objetos_por_categoria'}), name='objetos_por_categoria'),
    path('generar_excel/', views.generar_excel, name='generar_excel'),
    path('descargar_excel/', views.descargar_excel, name='descargar_excel'),
    path('', include(router.urls)),
]