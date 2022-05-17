from setuptools import setup
setup(
    name='auto-location-mapping',
    version='1.2.4',
    author='Derrick Alvarez',
    author_email='derrickcanbereached@gmail.com',
    install_requires=['pandas','pyautogui', 'openpyxl','progress'],
    entry_points='''
        [console_scripts]
        automap=automap:main
    ''',
)
