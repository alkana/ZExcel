# ZExcel
The goal of ZExcel is to rewrite PHPExcel with the Zephir language (http://zephir-lang.com/index.html) to create an php extension.

# Status
- 2016-03-30: Partial reader xlsx implemented (sheet, values and calculated values)
- 2016-02-28: The library is always not usable but lot of classes has been implemented
- 2015-05-12: The library is not usable at all.

| Types        | Readers | Writers |
| ------------ | ------- | ------- |
| csv          | none    | none    |
| excel2003xml | none    | none    |
| excel2007    | partial | none    |
| excel5       | none    | none    |
| gnumeric     | none    | none    |
| html         | none    | none    |
| oocalc       | none    | none    |
| sylk         | none    | none    |

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
