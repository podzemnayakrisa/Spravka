using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

public class ResponseItem : INotifyPropertyChanged
{

    public event PropertyChangedEventHandler PropertyChanged;

    private string _fullName = "";
    private string _email = "";
    private DateTime _requestDate = DateTime.Now;
    private string _course;
    private string _educationForm;
    private string _basis;
    private string _status = "Новый";
    private bool _isReady;


    public string FullName
    {
        get => _fullName;
        set => SetField(ref _fullName, value ?? "");
    }

    public string Email
    {
        get => _email;
        set => SetField(ref _email, value ?? "");
    }

    public DateTime RequestDate
    {
        get => _requestDate;
        set => SetField(ref _requestDate, value);
    }

    public string Course
    {
        get => string.IsNullOrWhiteSpace(_course) ? "Не указано" : _course;
        set => SetField(ref _course, value);
    }

    public string EducationForm
    {
        get => string.IsNullOrWhiteSpace(_educationForm) ? "Не указано" : _educationForm;
        set => SetField(ref _educationForm, value);
    }

    public string Basis
    {
        get => string.IsNullOrWhiteSpace(_basis) ? "Не указано" : _basis;
        set => SetField(ref _basis, value);
    }

    public string Status
    {
        get => _status;
        set
        {
            string newValue;
            if (value == "Готово")
                newValue = "Готово";
            else if (value == "В работе")
                newValue = "В работе";
            else
                newValue = "Новый";

            if (_status != newValue)
            {
                _status = newValue;
                _isReady = _status == "Готово";
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsReady));
            }
        }
    }

    public bool IsReady
    {
        get => _isReady;
        set
        {
            if (_isReady != value)
            {
                _isReady = value;
                Status = value ? "Готово" : "В работе";
                OnPropertyChanged();
            }
        }
    }
    private string _group;
    public string Group
    {
        get => _group;
        set
        {
            if (_group != value)
            {
                _group = value;
                OnPropertyChanged(nameof(Group));
            }
        }
    }
    private string _certificateNumber;
    public string CertificateNumber
    {
        get => _certificateNumber;
        set => SetField(ref _certificateNumber, value);
    }
    protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return false;

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

}