# ZExcel
The goal of ZExcel is to rewrite PHPExcel with the Zephir language (http://zephir-lang.com/index.html) to create an php extension.

# Status
- 2015-12-05: The library is not useable at all.
- 2016-02-28: The library always not usable but lot of classes has been implemented (~80%, you can easily test it with phpunit) 

# How to install ?
1. Install zephir under your server/vagrant/docker (@see https://github.com/phalcon/zephir#readme)
2. Clone ZExcel project (git clone https://github.com/alkana/ZExcel.git)
3. Execute zephir to compile the extension (zephir build)
4. Add the extension to your php.ini (or add new file on module-available - zexcel.ini -)
```
  [ZExcel]
  extension=zexcel.(so|dll)
``` 
5. check if the module is active
```
php -m | grep zexcel &> /dev/null && echo 'ZExcel is active !'
```

# How to test ?
When the extension has been installed, you can easily check the advencement with phpunit:
```
cd {project/folder} && phpunit -c unitTests/
```

# Usage
@TODO
