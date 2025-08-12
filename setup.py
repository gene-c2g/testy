from distutils.core import setup
from setuptools.command.install import install
import atexit

def _post_install():
    import nltk
    print('POST INSTALL')
    nltk.download("punkt")
    nltk.download("punkt_tab")

class PostInstallCommand(install):
    """Post-installation for installation mode."""
    def run(self):
        install.run(self)
        # PUT YOUR POST-INSTALL SCRIPT HERE or CALL A FUNCTION
        atexit.register(_post_install)

setup(name='testy',
      version='1.0.7',
      description='C2G Test Suite Helper Tool',
      author = 'Charlie Lenahan',
      author_email = 'clenahan@cloud2gnd.com',
      install_requires= ['python-docx', 'openpyxl', 'nltk'],
      py_modules=['testy', 'tcrl', 'iopimport', 'tca', 'wordhelper','excelhelper','bcolors'],
       cmdclass={
        'install': PostInstallCommand,
       },
      )