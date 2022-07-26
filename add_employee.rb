require 'watir'

b = Watir::Browser.new

b.goto'localhost/orangehrm'

b.driver.manage.window.maximize

b.text_field(name:'username').set('admin')

b.text_field(name:'password').set('Admin@123')

b.button(type:'submit').click

b.element(text:'Add Employee').click

b.text_field(name:'First Name').set('Phuc')

sleep(1)

b.text_field(name:'Middle Name').set('Thien')
sleep(1)

b.text_field(name:'Last Name').set('Tran')
sleep(1)
b.span(class:"oxd-switch-input oxd-switch-input--active --label-right").click
sleep(1)
b.text_field(autocomplete:'off').set('user2')
sleep(1)
tfs = b.text_fields(type:'password', autocomplete:'off')
tfs.each do |div|
  div.set('User@123')
end

b.button(type:'submit').click

sleep(4)
