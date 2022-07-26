# import lib
require 'watir'
require 'win32ole'

# connect xlsx
wb = WIN32OLE.connect('I:\Project\TestOrangeHrm\testcase.xlsx')
ws = wb.Worksheets('Sheet1')

stop = false
i = 2

while stop == false

  # connect browser
  b = Watir::Browser.new

  b.goto('localhost/orangehrm')

  b.driver.manage.window.maximize

  # login
  b.text_field(name:'username').set(ws.Cells(i,'B').value)

  sleep(1)

  b.text_field(name:'password').set(ws.Cells(i+1,'B').value)

  sleep(1)

  b.button(type:'submit').click

  # go to page assign leave
  b.element(href:'/orangehrm/web/index.php/leave/viewLeaveModule').click

  b.li(text:'Assign Leave').click

  # start test
  if i != 9
    b.text_field(placeholder:'Type for hints...').set(ws.Cells(i+2,'B').value)

    sleep(2)

    b.element(role:'listbox').click

    ws.Cells(i+2,'D').value = b.text_field(placeholder:'Type for hints...')

    sleep(1)
  end

  if i != 16
    b.element(text:'-- Select --').click

    sleep(1)

    b.element(text: ws.Cells(i+3,'B').value).click

    ws.Cells(i+3,'D').value = b.element(text:'Go give birth').text

    sleep(1)
  end

  if i != 23
    b.element(class:'oxd-icon bi-calendar oxd-date-input-icon').click

    sleep(1)

    b.element(text:'Today').click

    sleep(1)
  end

  b.element(class:"oxd-textarea oxd-textarea--active oxd-textarea--resize-vertical").set(ws.Cells(i+4,'B').value)

  ws.Cells(i+4,'D').value = b.textarea(value:'Allowed leave')

  sleep(1)

  b.button(type:'submit').click

  sleep(1)

  if i == 9

    ws.Cells(i+2,'D').value = b.element(text:'Required').text

  end

  if i == 16
    ws.Cells(i+3,'D').value = b.element(text:'Required').text
  end

  if i == 23
    ws.Cells(i+5,'D').value = b.element(text:'Required').text
  end

  if i == 2

    b.element(text:'Ok').click

    sleep(1)

    ws.Cells(i+5,'D').value = b.p(text:'Successfully Saved').text

    sleep(1)

  end

  i = i + 7

  if i > 29

    stop = true

  end

  b.close

end
