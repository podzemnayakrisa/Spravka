using System;
using System.ComponentModel;

namespace Spravka.Models
{
    public class GroupItem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _name;
        private DateTime _startDate;
        private DateTime _endDate;
        private string _certificateTemplateId;
        private string _directorName;

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        public DateTime StartDate
        {
            get => _startDate;
            set
            {
                if (_startDate != value)
                {
                    _startDate = value;
                    OnPropertyChanged(nameof(StartDate));
                }
            }
        }

        public DateTime EndDate
        {
            get => _endDate;
            set
            {
                if (_endDate != value)
                {
                    _endDate = value;
                    OnPropertyChanged(nameof(EndDate));
                }
            }
        }

        public string CertificateTemplateId
        {
            get => _certificateTemplateId;
            set
            {
                if (_certificateTemplateId != value)
                {
                    _certificateTemplateId = value;
                    OnPropertyChanged(nameof(CertificateTemplateId));
                }
            }
        }

        public string DirectorName
        {
            get => _directorName;
            set
            {
                if (_directorName != value)
                {
                    _directorName = value;
                    OnPropertyChanged(nameof(DirectorName));
                }
            }
        }

        // Форматированное представление периода обучения
        public string StudyPeriod => $"{StartDate:dd.MM.yyyy} - {EndDate:dd.MM.yyyy}";

        // Длительность обучения в месяцах
        public int DurationMonths => (EndDate.Year - StartDate.Year) * 12 + EndDate.Month - StartDate.Month;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // Метод для проверки активности группы на текущую дату
        public bool IsActiveGroup(DateTime? date = null)
        {
            var checkDate = date ?? DateTime.Now;
            return checkDate >= StartDate && checkDate <= EndDate;
        }

        // Метод для создания копии объекта
        public GroupItem Clone()
        {
            return new GroupItem
            {
                Name = this.Name,
                StartDate = this.StartDate,
                EndDate = this.EndDate,
                CertificateTemplateId = this.CertificateTemplateId,
                DirectorName = this.DirectorName
            };
        }
    }
}