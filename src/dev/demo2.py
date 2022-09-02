from selenium import webdriver

# https://xmadoka.hatenablog.com/entry/2020/03/07/023200
driver = webdriver.Chrome('/Users/masataka/Coding/Pythons/Licosha/Display/assets/chromedriver')
driver.get("https://forms.app/recordimage/6308f6c70169057963106ed8/435a4a43-255d-11ed-8d27-4201c0a80002")
src = driver.find_element_by_tag_name("img").get_attribute("src")
print(src)