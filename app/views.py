from django.http import HttpResponseRedirect
from django.shortcuts import redirect, render
from django.core.mail import send_mail
from django.conf import settings
from django.contrib import messages
# Create your views here.
def home(request):  # sourcery skip: extract-duplicate-method, extract-method, remove-redundant-if
    if request.method == 'POST':
        Name= request.POST['Name']
        Email = request.POST['Email']
        Phone_Number = request.POST['Phone_Number']
        Message = request.POST['Message']
        message = f"Name: {Name} Email: {Email} Phone_Number: {Phone_Number} Message: {Message}"
        send_mail(
            'Contact Form',
            message,
            'settings.EMAIL_HOST_USER',
            [Email],
            fail_silently=True,
        )
        messages.success(request, 'Successfully Sent The Message!')
        return redirect('/')
    return render(request, 'index.html')