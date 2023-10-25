build:
    vagrant up
    vagrant ssh-config > .vagrant-ssh-config
    ssh -F .vagrant-ssh-config default powershell C:/vagrant/Elitepad900.ps1
