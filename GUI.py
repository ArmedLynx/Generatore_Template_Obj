import kivy
from kivy.app import App

from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
  
class MyFirstKivyApp(App):
      
    def build(self):
        layout = GridLayout(rows=2, cols=2)
        label1 = Label(text="[b]Hello1[/b]", markup=True)
        label2 = Label(text="[i]Hello2[/i]", markup=True)
        label3 = Label(text="[u]Hello3[/u]", markup=True)
        label4 = Label(text="[s]Hello4[/s]", markup=True)
        layout.add_widget(label1)
        layout.add_widget(label2)
        layout.add_widget(label3)
        layout.add_widget(label4)
        return layout          
  
MyFirstKivyApp().run()          