1. el fornt se subio a un bucket de s3
2. se creo http api
3. se crearon dependencias en wsl, en el proyecto back para subirse :

mkdir -p lambda_layer_linux_312/python/lib/python3.12/site-packages

pip3.12 install python-docx mammoth lxml -t lambda_layer_linux_312/python/lib/python3.12/site-packages

cd lambda_layer_linux_312
zip -r layer.zip python

