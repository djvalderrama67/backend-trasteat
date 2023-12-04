from django.db import models
from django.db.models import Sum

# Create your models here.

class Categoria (models.Model):
    idCategoria = models.AutoField(primary_key=True, verbose_name='ID')
    nombre = models.CharField(max_length=50, verbose_name='Nombre de la categoria')

    def __str__(self):
        return self.nombre

    class Meta:
        verbose_name = 'Categoria'
        verbose_name_plural = 'Categorias'


class Vehiculo (models.Model):
    idVehiculo = models.AutoField(primary_key=True, verbose_name='ID')
    nombre = models.CharField(max_length=100, verbose_name='Nombre del vehiculo')
    capacidad_min = models.IntegerField(verbose_name='Capacidad minima del vehiculo',default=1)
    capacidad_max = models.IntegerField(verbose_name='Capacidad maxima del vehiculo',default=60)
    volumen_carga_min = models.IntegerField(verbose_name='Volumen minimo de carga (m³)', default=2)
    volumen_carga_max = models.IntegerField(verbose_name='Volumen maximo de carga (m³)', default=120)

    def __str__(self):
        return self.nombre

    class Meta:
        verbose_name = 'Vehiculo'
        verbose_name_plural = 'Vehiculos'

class Bodega (models.Model):
    idBodega = models.AutoField(primary_key=True, verbose_name='ID')
    nombre = models.CharField(max_length=50, verbose_name='Nombre de la bodega')
    area = models.FloatField(verbose_name='Area (m²)', blank=True, null=True)
    volumen = models.FloatField(verbose_name='Volumen (m³)', blank=True, null=True)
    altura = models.FloatField(verbose_name='Altura (m)')
    largo = models.FloatField(verbose_name='Largo (m)')
    ancho = models.FloatField(verbose_name='Ancho (m)')

    def save(self, *args, **kwargs):
        self.area = round(self.largo * self.ancho, 2)
        self.volumen = round(self.area * self.altura, 2)
        super(Bodega, self).save(*args, **kwargs)

    def __str__(self):
        return self.nombre

    class Meta:
        verbose_name = 'Bodega'
        verbose_name_plural = 'Bodegas'

class Objeto (models.Model):
    idObjeto = models.AutoField(primary_key=True, verbose_name='ID')
    nombre = models.CharField(max_length=50, verbose_name='Nombre del objeto')
    categoria = models.ForeignKey(Categoria, on_delete=models.CASCADE, verbose_name='Categoria')
    largo = models.FloatField(verbose_name='Largo (m)')
    ancho = models.FloatField(verbose_name='Ancho (m)')
    alto = models.FloatField(verbose_name='Alto (m)')
    volumen = models.FloatField(verbose_name='Volumen (m³)', blank=True, null=True)

    def save(self, *args, **kwargs):
        self.volumen = round(self.largo * self.ancho * self.alto, 2)
        super(Objeto, self).save(*args, **kwargs)

    def __str__(self):
        return self.nombre

    def nombre_categoria(self):
        return self.categoria.nombre

    class Meta:
        verbose_name = 'Objeto'
        verbose_name_plural = 'Objetos'

class Calculo(models.Model):
    idCalculo = models.AutoField(primary_key=True, verbose_name='ID')
    objetos = models.ManyToManyField(Objeto, verbose_name='Objetos seleccionados')
    volumen_Total = models.FloatField(verbose_name='Volumen Total (m³)')
    idVehiculo = models.ForeignKey(Vehiculo, on_delete=models.CASCADE)
    idBodega = models.ForeignKey(Bodega, on_delete=models.CASCADE)

    def save(self, *args, **kwargs):
        volumen_total = self.objetos.aggregate(Sum('volumen'))['volumen__sum'] or 0.0
        self.volumen_Total = volumen_total
        super(Calculo, self).save(*args, **kwargs)

    def __str__(self):
        return f'Calculo {self.idCalculo}: {", ".join([str(objeto) for objeto in self.objetos.all()])}'

    class Meta:
        verbose_name = 'Calculo'
        verbose_name_plural = 'Calculos'












