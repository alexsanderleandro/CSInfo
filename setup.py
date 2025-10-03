from setuptools import setup, find_packages

setup(
    name="csinfo",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        'wmi',
        'python-dotenv',
        'requests',
        'psutil',
        'pywin32; platform_system=="Windows"',
    ],
    python_requires='>=3.6',
    author="Seu Nome",
    author_email="seu.email@exemplo.com",
    description="Pacote para coleta de informações do sistema",
    keywords="system info hardware",
    entry_points={
        'console_scripts': [
            'csinfo=csinfo:main',
        ],
    },
    include_package_data=True,
)
