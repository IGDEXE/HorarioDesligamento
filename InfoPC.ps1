# Gerenciar horario de funcionamento do PC
# Ivo Dias

# Verificar LOGs do sistema
function Verificar-DesligamentoPC {
    param (
        [parameter(position=0)]
        $Computadores = $Env:COMPUTERNAME,
        [parameter(position=1)]
        $LogPath = "C:\TI\ValidarHorario\LOGs"
    )
    # Cria o hash
    $hash = Get-Date -Format yyyyMMddTHHmmssffff
    # Faz o loop para os computadores
    foreach ($Computador in $Computadores) {
        # Testa a conexao
        Write-Host "Verificando o computador $Computador"
        if (Test-Connection -ComputerName $Computador -Quiet) {
            try {
                # Cria a pasta de LOGs
                $valida = Test-Path "$LogPath"
                if ($valida -ne "True") {
                    # Cria a pasta
                    $off = mkdir $LogPath
                    Write-Host "Criado a pasta de LOGs: $LogPath"
                }
                # Recebe os dados do desligamento
                $info = Get-WinEvent -ComputerName $Computador -FilterHashtable @{logname='System'; id=1074}  | ForEach-Object {
                    $rv = New-Object PSObject | Select-Object Date, User, Action, Process, Reason, ReasonCode, Comment
                    $rv.Date = $_.TimeCreated
                    $rv.User = $_.Properties[6].Value
                    $rv.Process = $_.Properties[0].Value
                    $rv.Action = $_.Properties[4].Value
                    $rv.Reason = $_.Properties[2].Value
                    $rv.ReasonCode = $_.Properties[3].Value
                    $rv.Comment = $_.Properties[5].Value
                    $rv
                } | Select-Object Date, Action, User
                $info = $info[0].Date
                # Recebe os dados da inicializacao
                $ini = Get-CimInstance -ClassName win32_operatingsystem
                $ini = $ini.lastbootuptime
                # Grava os dados
                Add-Content -Path "$LogPath\Relatorio-$hash.log" -Value $Computador
                Add-Content -Path "$LogPath\Relatorio-$hash.log" -Value "|Ultima inicializacao:$ini|"
                Add-Content -Path "$LogPath\Relatorio-$hash.log" -Value "|Ultimo desligamento :$info|"
                Add-Content -Path "$LogPath\Relatorio-$hash.log" -Value "__"
                # Mostra os dados
                Write-Host "As informacoes do $Computador foram adicionados ao arquivo: $LogPath\Relatorio-$hash.log"
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-Host "Erro ao verificar o computador $Computador"
                Write-Host "Erro: $ErrorMessage"
            }
        }
        else {
            # Mostra mensagem de erro
            Write-Host "xx Erro ao acessar o computador $Computador"
            Add-Content -Path "$LogPath\OFF-$hash.log" -Value $Computador
        }
    }
}

# Faz a interface com o usuario
# Configuracoes
# Recebe a lista de Computadores
$Computadores = get-content "C:\TI\ValidarHorario\ListaComputadores.txt"
Write-Host "Aguarde o carregamento"
Verificar-DesligamentoPC $Computadores
Write-Host "Procedimento concluido"
pause