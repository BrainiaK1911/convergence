from ._anvil_designer import MainTemplate
from anvil import *
import anvil.server
import anvil.tables as tables
import anvil.tables.query as q
from anvil.tables import app_tables
import anvil.media
import validation
import time

class Main(MainTemplate):
  def __init__(self, **properties):
    # Set Form properties and Data Bindings.
    self.init_components(**properties)
    # Any code you write here will run when the form opens.
    # anvil.server.call("say_hello", self.name_box.text
        


  def primary_color_1_click(self, **event_args):
    """This method is called when the button is clicked"""
    ownership_list = self.ownership.text.split(',')
    for ownership in ownership_list:
      m = anvil.server.call("get_output", self.spdr_file_uploaded.text, self.mspr_file_uploaded.text, self.spdr_file.file, self.mspr_file.file, self.territory.text, ownership, self.email.text)
      download(m)
      time.sleep(7)
    

  def territories_pressed_enter(self, **event_args):
    """This method is called when the user presses Enter in this text box"""

  def email_pressed_enter(self, **event_args):
    """This method is called when the user presses Enter in this text box"""
    pass

  def ownership_pressed_enter(self, **event_args):
    """This method is called when the user presses Enter in this text box"""
    pass

  def mspr_file_change(self, file, **event_args):
    """This method is called when a new file is loaded into this FileLoader"""
    self.mspr_file_uploaded.text = file.name

  def spdr_file_change(self, file, **event_args):
    """This method is called when a new file is loaded into this FileLoader"""
    self.spdr_file_uploaded.text = file.name
