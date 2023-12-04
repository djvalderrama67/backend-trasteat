from rest_framework import serializers
from .models import *

class CategoriaSerializer(serializers.ModelSerializer):
    class Meta:
        model = Categoria
        fields = '__all__'

class VehiculoSerializer(serializers.ModelSerializer):
    class Meta:
        model = Vehiculo
        fields = '__all__'

class BodegaSerializer(serializers.ModelSerializer):
    class Meta:
        model = Bodega
        fields = '__all__'

class ObjetoSerializer(serializers.ModelSerializer):
    categoria = serializers.StringRelatedField()
    class Meta:
        model = Objeto
        fields = '__all__'


class CalculoSerializer(serializers.ModelSerializer):
    class Meta:
        model = Calculo
        fields = '__all__'

