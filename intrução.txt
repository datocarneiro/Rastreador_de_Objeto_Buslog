# Para ativar um ambiente virtual Seguir as etapas a baixo


# 1 - Verificar se o python esta intstalado
#       python --version


# 2 - criar o ambiente 
#       python -m venv venv (o segunfo venv é o nome do ambiente, pode ser qualquer nome)


# 3 - ativar o ambente virtual 
        #   - precisa estar no diretório dp projeto 
        #   - para verificar em qual pasta estamos usar o comando no terminal "pwd"

        #   - POWERSHELL - 
        #           no power sheel o comando para ativar é: .\venv\Scripts\Activate.ps1
        #           se der erro precisa abrir o powershell nativo e executar o somando:
        #                Set-ExecutionPolicy -Scape Currentuser -ExecutionPolicy RemoteSigned


        #   - CMD - 
        #           no cmd padrãol o comando para ativar é: .\env\Scripts\Activate.bat

        #   - BASH - 
        #           - no bash o comando é: `source ./venv/Scripts/activate`

# 4 - Desativar o ambiente virtual (somente)
#          comando:deactivate


# REQUIREMENTS.TXT 
#           Para criar o arquivo use o comando: pip freeze > requirements.txt