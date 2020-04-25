class LoginFacebook():

    def __init__(self,driver):
        self.driver = driver

    def findEmail(self):
        return self.driver.find_element_by_id("email")

    def setEmail(self,email):
        self.findEmail().send_keys(email)

    def findPassword(self):
        return self.driver.find_element_by_id("pass")

    def setPassword(self,password):
        self.findPassword().send_keys(password)

    def login(self):
        self.driver.find_element_by_id("u_0_b").click()