   M S B A M L   `   `   `  ����T  PMathematicalComputingTool, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullE'clr-namespace:MathematicalComputingToolMathematicalComputingTool  5   	   *    $MathematicalComputingTool.MainWindow   5   	   \ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   9local'clr-namespace:MathematicalComputingTool     5   	     � Title$  MathematicalComputingTool��5   	   $	��150��6+   $	��530��68    � 
ResizeMode$ NoResize=�5	   	    � WindowStartupLocation$ CenterScreen=�5
   	   .��5      �  � RowDefinitions 5   
   �� 5      $��* �6   �� 5      $��* �6   �� 5      $��2* �6    � ColumnDefinitions 5   
   �� 5      $
��auto �6   �� 5      $
��auto �6   .+�5   
   �� -   5        $��Txtファイル選択��5      $��0��6G   $��0��6T   $��5q�6d   �� 5   
   $��0��68   $��1��6E   $��5q�6U   $��Center=�6`   $	��400��6{   $
��True��6�   ��6   ��  �� Path$ SelectedTxtPath�6�   �� 5   
   -   5        $��Excelファイル選択��5      $��1��6K   $��0��6X   $��5q�6h   �� 5!   
   $��1��6:   $��1��6G   $��5q�6W   $��Center=�6b   $	��400��6}   $
��True��6�   ��6   �� $ SelectedExcelPath�6�   �� 5$   
   $��2��6   $��0��6   $��Center=�6-   $��Center=�6O   .��6s   結果6y   � 5&   
   $��2��6   $��1��6   $��5q�6,   .+�5'      �� .��5(      �� $��Center=�6<   $	��400��6X    �� TextWrapping$
 Wrap=�6d    �� 
IsReadOnly$
 True��6x   ��6   �� $ CreatedPath�6x   5)      5*      5,      5-      