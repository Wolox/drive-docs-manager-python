drive-docs-manager-python
===============

This script was designed to (parcially) automize the performance evaluation process

## First Steps

### Installing pygsheets
```pip3 install git+git://github.com/nithinmurali/pygsheets@75594dc58a5a9671edea2283369bb190aac36fb3```

### Installing pydrive
```pip3 install pydrive```

### Installing readchar
```pip3 install readchar```

### Uninstall & install google-auth-oauthlib in this version (0.3.0)
When installing pygsheets, it automatically installed google-auth-oauthlib in the latest version, but it don't work (0.4.0)
It throws the error: `oauthlib.oauth2.rfc6749.errors.InvalidGrantError: (invalid_grant) code_verifier or verifier is not needed.`

To solve this, install the version 0.3.0
```pip3 install google-auth-oauthlib==0.3.0```

### For error "ImportError: No module named oauth2client.file"
`pip3 install --upgrade oauth2client`

### Installing Percol
```pip3 install Percol```

## Documentation

### Readchar
> Documentation for readchar: https://github.com/magmax/python-readchar

### Percol
> Documentation for percol: https://github.com/mooz/percol

### PyDrive
> Documentation for PyDrive: https://pythonhosted.org/PyDrive/

### To create credentials in pygsheets
> https://pygsheets.readthedocs.io/en/stable/authorization.html

### To create credentials in pydrive
> https://pythonhosted.org/PyDrive/quickstart.html

You have to download the credentials.json logging with gauth.LocalWebserverAuth()

## About

This project is maintained by [Wolox](https://github.com/Wolox) and it was written by [Wolox](http://www.wolox.com.ar).

![Wolox](https://raw.githubusercontent.com/Wolox/press-kit/master/logos/logo_banner.png)