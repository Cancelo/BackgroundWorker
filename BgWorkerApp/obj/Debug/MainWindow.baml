   M S B A M L   `   `   `  ����F  BBgWorkerApp, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null)clr-namespace:BgWorkerAppBgWorkerApp  5   	       BgWorkerApp.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   +localclr-namespace:BgWorkerApp     5   	     � Title$  MF Reporter��5   
   $	��450��6   $	��700��6+    � 
ResizeMode$ NoResize=�67    � WindowStyle$
 None=�6M   
�������$��1q�6w    � 	Resources 5	       �� MergedDictionaries �� Source�� 5
       5   
   �� 5      $ 
Style.xaml?�6!   %    5      .��5      ]� .+�5   
   �� -   5        $	��Top=�5      $��50��6*   ���#0f3463.+�5      ]� $��5q�6   .+�5      �� -   5           CloseBtn#��[  5      $��Images/close.png��6:   $��40��6T   $��Right=�6_   6v   �� 5      $��Images/logoLm.png��6   $	��200��63   $
��Left=�6?   $
��Left=�6U   $��5,0,0,0q�6p   5      5      ]� 5   
   $	��Top=�6   $��
10,10,10,0q�6)   .+�5      (  "System.Windows.Controls.DatePicker  -   5          Name$ startDatePicker��5         FirstDayOfWeek$ Monday=�62   �� 5      $��-��6     5      -   5        $ endDatePicker��5      �� 5      -   5        $"��btnDoAsynchronousCalculation��5       	 Boton#��[ 69   $��Right=�6�   $��3q�6�   $��Right=�6�   {�.  .��6�   	Empezar6�   �� 5      -   5        $��btnCreateReport��5      #��[ 6,   $��Right=�6K   $��Right=�6b   {�.  $��3q�6�   $	��5,0q�6�   .��6�   Generar informe6�   5      ]� 5   
   $��Bottom=�6   .+�5       �� $��Bottom=�6   ��5!      �� 5"      .R�5#      �  � ColumnDefinitions 5$      �� 5%   "   $	��100 �63   �� 5&   "   $
��Auto �63   �� 5'   "   $��* �63   �� 5(   "   $
��Auto �63   �� 5)   "   $	��250 �63   5+      5,      .@�5.      �� .��5/      o� $	��Top=�6   .@�50      m� $��_Configuración��6#   .@�51      m� -   5        $��_Seleccionar archivo...��51   '   6H   m� 52      $��_Opciones de lectura��6'   .@�53   "   ]� $	��180��6,   .+�54   &   �� $
��Left=�6,   $��_Rango de columnas��6B   �� 55   &   -   5        	 ��Name$	 totalColsTxt��55   .   $��Right=�6D   $��30��6[   $	��269��6f   $��Center=�6q   $��Right=�6�   
 � ToolTip$
 Inicio��6�   �� 56   &   -	   5        $	 startColTxt��56   .   $��Right=�6C   $��30��6Z   $��7��6e   $��Center=�6n   $��Right=�6�   $
 Final��6�   57   #   58      59      5:      5;      �� 5<      $��1��6   �� 5=      $��2��6    .��5>      �� -
   5         ��Name$ lblPath��5>       
�������61   5?      �� 5@      $��3��6   �� 5A      $��4��6    .��5B      � -   5         �Name$ pbCalculationProgress��5B   G     ProgressBar#��[ 6"   $	��237��6f   $��15��6r   5C      5D      5E      &   System.Windows.Controls.DataGrid  5G   
   -   5        $��	lbResults��5G      $��
10,10,10,0q�6$   6:   5H      5I      