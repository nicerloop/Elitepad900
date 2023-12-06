build:
    vagrant up
    vagrant snapshot restore clean || vagrant snapshot save clean
    vagrant ssh -- powershell C:/vagrant/Elitepad900.ps1

clean:
    -vagrant destroy --force
