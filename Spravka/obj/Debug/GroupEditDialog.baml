   M S B A M L   `   `   `  ����  Spravka    Spravka.GroupEditDialog   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	     � Title$/  )Редактирование группы��5   	   $	��250��5   	   $	��400��6    � WindowStartupLocation$ CenterOwner=�5	   	   .��5
      � $��10q�6    � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   �� $$��Название группы:��6   $��0��6,   �� 5   
   -   5         ��Name$ txtName��5      $��1��6#   $��0,0,0,10q�60   �� 5   
   $-��'Дата начала обучения:��6   $��2��61   (  "System.Windows.Controls.DatePicker  5   
   -   5          Name$ dpStartDate��5      $��3��6*   $��0,0,0,10q�67   �� 5   
   $3��-Дата окончания обучения:��6   $��4��64     5   
   -   5        $ 	dpEndDate��5      $��5��6(   $��0,0,0,20q�65   �� 5   
   $��6��6    �� Orientation$ 
Horizontal=�6"   $��Right=�6;   .+�5      �� -   5        $
��ОК��5      $��80��6"   $��0,0,10,0q�6-   6?   �� 5       -   5        $��Отмена��5       $��80��6&   61   5!      5"      5#      