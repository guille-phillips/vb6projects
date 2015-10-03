function ToggleDisplay(elementId)
{
  // Get a reference to an element in the page and switch its display property.
  var theElement = document.getElementById(elementId);
  if(theElement.style.display == 'none')
  {
      theElement.style.display = 'block';
  }
  else
  {
    theElement.style.display = 'none';
  }
}