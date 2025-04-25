from django.shortcuts import render
from django.shortcuts import HttpResponse

# Create your views here.
def neglavnaya(request):
    return render(request, 'spravka/index.html')

def glavnaya(request):
    return HttpResponse('прощайте миньенчики')

def register(request):
    return render(request, 'spravka/register.html')





