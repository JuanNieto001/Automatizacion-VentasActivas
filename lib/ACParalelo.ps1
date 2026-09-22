# =====================================================================
#  ACParalelo.ps1 - Consulta de estado con varias instancias de AC a la vez
#
#  Solo aplica a la pasada de ESTADO (activo / no activo), que se maneja
#  integramente por mensajes de Windows: no usa raton ni teclado, funciona
#  con AC en segundo plano y por eso varias instancias no se estorban.
#
#  La pasada del HISTORIAL no se puede paralelizar: abrir la ficha del
#  cliente exige un clic con el raton real, y el raton es unico por sesion
#  de Windows.
# =====================================================================

. "$PSScriptRoot\Win32.ps1"
. "$PSScriptRoot\Remote.ps1"
. "$PSScriptRoot\AC.ps1"

function Get-ACControlesFijos {
    <#  Captura una sola vez los controles que no cambian durante la sesion.
        Redescubrirlos en cada sondeo costaria mas que la consulta misma.  #>
    param([Parameter(Mandatory=$true)][int]$ProcessId)

    $c = Get-ACContext -ProcessId $ProcessId
    if (-not $c -or -not $c.BotonBuscar) { return $null }
    return [pscustomobject]@{
        ProcessId  = $ProcessId
        Principal  = $c.Principal
        MDIClient  = $c.MDIClient
        Edit       = $c.EditCriterio
        Radio      = $c.RadioMin
        Buscar     = $c.BotonBuscar
    }
}

function Invoke-ACClicActivado {
    <#  Envia un clic por mensaje a un control de AC, activando antes su ventana.

        Hallazgo medido: AC solo atiende el clic enviado por mensaje si esa
        instancia es la ventana en primer plano en ese instante. Ademas, el
        primer clic sobre el panel actua de "cebador" (deja el foco dentro del
        panel) y el que cuenta es el siguiente.

        Esto NO impide paralelizar: el primer plano solo hace falta durante el
        clic (unas decimas), mientras que la espera de Oracle y la lectura de
        las grillas funcionan con la ventana de fondo. Las instancias se turnan
        el primer plano y sus esperas se solapan.  #>
    param(
        [Parameter(Mandatory=$true)]$Fijos,
        [Parameter(Mandatory=$true)][IntPtr]$Control,
        [IntPtr]$Cebo = [IntPtr]::Zero,
        [int]$X = 10, [int]$Y = 10
    )
    if (-not (Set-WindowFocus -Handle $Fijos.Principal.Handle)) { return $false }
    if ($Cebo -ne [IntPtr]::Zero) {
        Invoke-PostClick -Handle $Cebo
        Start-Sleep -Milliseconds 120
    }
    Invoke-PostClick -Handle $Control -X $X -Y $Y
    # El clic va POSTEADO: queda en la cola de AC. Hay que darle tiempo a
    # procesarlo MIENTRAS esta ventana sigue en primer plano, porque si otra
    # instancia se lo roba antes, AC lo descarta y la consulta se pierde.
    Start-Sleep -Milliseconds 400
    return $true
}

function Get-VentanaResultados {
    <#  Sondeo barato: solo mira los hijos del cliente MDI.  #>
    param([Parameter(Mandatory=$true)]$Fijos)
    if (-not $Fijos.MDIClient) { return $null }
    return (Get-ChildHandles -Parent $Fijos.MDIClient.Handle |
            Where-Object { $_.Visible -and $_.Class -eq 'ThunderRT6FormDC' -and $_.Text -like 'Resultados*' } |
            Select-Object -First 1)
}

function Start-ACInstancias {
    <#  Levanta N instancias de AC con sesion iniciada.  #>
    param(
        [Parameter(Mandatory=$true)][int]$Cantidad,
        [Parameter(Mandatory=$true)][string]$Password,
        [string]$BaseDatos = 'AC_PRODUCCION'
    )

    Stop-ACSession -Todas
    $inst = New-Object System.Collections.ArrayList

    for ($i = 1; $i -le $Cantidad; $i++) {
        try {
            $ctx = Start-ACInstancia -Password $Password -BaseDatos $BaseDatos -Silencioso
            $fijos = Get-ACControlesFijos -ProcessId $ctx.ProcessId
            if (-not $fijos) { throw "la instancia no expuso el panel de criterios" }
            [void]$inst.Add([ordered]@{
                Indice   = $i
                Fijos    = $fijos
                Estado   = 'libre'
                Numero   = $null
                T0       = $null
                Intentos = 0
                UltClic  = [datetime]::MinValue
                UltBuscar= [datetime]::MinValue
                Reenvios = 0
                Consultas= 0
                Carga    = 0
                T0Leyendo= [datetime]::MinValue
                Ctrls    = $null
                Mudo     = 0
            })
            Write-Paso "Instancia $i de $Cantidad lista (PID $($ctx.ProcessId))" "OK"
        } catch {
            Write-Paso "No se pudo levantar la instancia ${i}: $($_.Exception.Message)" "ERROR"
        }
    }

    if ($inst.Count -eq 0) { throw "No se pudo iniciar ninguna instancia de AC." }
    return $inst
}

function Invoke-ACConsultaMasiva {
    <#  Consulta el estado de una lista de numeros repartiendolos entre las
        instancias disponibles. Devuelve un resultado por numero, en el orden
        en que fueron terminando, mas estadisticas de tiempo.

        -OnResultado recibe cada resultado apenas esta listo, para que quien
        llame pueda ir mostrando el avance.  #>
    param(
        [Parameter(Mandatory=$true)][string[]]$Numeros,
        [Parameter(Mandatory=$true)]$Instancias,
        [string]$Password,
        [string]$BaseDatos = 'AC_PRODUCCION',
        [scriptblock]$OnResultado,
        [int]$TimeoutPorNumeroSeg = 75,
        # Plazo mas corto para el caso "la ventana abrio pero el arbol esta
        # vacio", que es como se resuelven los numeros que no existen en AC.
        # 0 lo desactiva y se vuelve al plazo completo. 45 s deja margen sobre
        # la consulta mas lenta que si encontro (47 s en total, y de esos solo
        # una parte con el arbol vacio).
        [int]$TimeoutVacioSeg = 45,
        [int]$IntervaloMs = 250,
        [int]$MaxReenvios = 4,
        # Se mide en BUSCAR enviados, no en numeros resueltos: ver el reciclado
        # preventivo mas abajo.
        [int]$ReciclarCada = 150,
        [int]$MaxRevivir = 3,
        # Sondeos seguidos sin respuesta antes de dar la instancia por colgada.
        # A 250 ms por vuelta y 2 s de plazo por sondeo, 5 son ~10 s de silencio.
        [int]$MaxSondeosMudos = 5,
        [switch]$Trace
    )

    $cola = New-Object System.Collections.Queue
    foreach ($n in $Numeros) { $cola.Enqueue($n) }

    $resultados = New-Object System.Collections.ArrayList
    $tiempos    = New-Object System.Collections.ArrayList
    $swGlobal   = [Diagnostics.Stopwatch]::StartNew()
    $rondasRevivir = 0
    $incompleto    = $false

    function Publicar {
        param($Inst, $Res)
        $seg = if ($Inst.T0) { ((Get-Date) - $Inst.T0).TotalSeconds } else { 0 }
        [void]$tiempos.Add($seg)
        $Res | Add-Member -NotePropertyName Segundos  -NotePropertyValue ([Math]::Round($seg, 2)) -Force
        $Res | Add-Member -NotePropertyName Instancia -NotePropertyValue $Inst.Indice -Force
        [void]$resultados.Add($Res)
        if ($OnResultado) { & $OnResultado $Res }
        # dejar la instancia lista para el siguiente numero
        Close-ACResultados -ProcessId $Inst.Fijos.ProcessId
        $Inst.Estado   = 'libre'
        $Inst.Consultas = $Inst.Consultas + 1
        $Inst.Numero   = $null
        $Inst.Intentos = 0
        # la ventana de resultados ya se cerro: sus handles no sirven para el
        # siguiente numero
        $Inst.Ctrls    = $null
        $Inst.Mudo     = 0
    }

    while ($cola.Count -gt 0 -or ($Instancias | Where-Object { $_.Estado -ne 'libre' })) {

        foreach ($inst in $Instancias) {

            if ($inst.Estado -eq 'muerta') { continue }

            # --- la instancia murio: se descarta para no colgar el ciclo ---
            if (-not (Get-ACProcess -ProcessId $inst.Fijos.ProcessId)) {
                if ($inst.Numero) {
                    Publicar -Inst $inst -Res ([pscustomobject]@{
                        Numero = $inst.Numero; Encontrado = $false
                        Mensaje = "La instancia de AC se cerro inesperadamente."; Filas = @()
                    })
                }
                $inst.Estado = 'muerta'
                Write-Paso "La instancia $($inst.Indice) se cerro; se continua con las demas." "WARN"
                continue
            }

            # --- reciclado preventivo ---
            # Medido en una corrida de 1712 numeros: tras ~400 consultas las
            # instancias se degradan y acaban cerrandose solas, y antes de
            # caerse empiezan a devolver "NO ENCONTRADO" falsos. Reiniciarlas
            # cada cierto numero de consultas evita ese deterioro.
            #
            # Lo que desgasta a AC son los BUSCAR enviados, no los numeros
            # resueltos: un numero que existe cuesta un BUSCAR, pero uno que no
            # aparece agota el timeout y cuesta 1 + MaxReenvios. Contando
            # numeros, un lote donde casi nada existe desgasta cinco veces mas
            # rapido de lo que el contador refleja y las instancias mueren
            # antes de alcanzar el umbral (visto el 21-sep: murieron a las ~28
            # consultas, con el umbral en 150). Por eso se cuenta la carga.
            if ($inst.Estado -eq 'libre' -and $ReciclarCada -gt 0 -and $Password -and
                $inst.Carga -ge $ReciclarCada -and $cola.Count -gt 0) {

                Write-Paso "Instancia $($inst.Indice): reciclando tras $($inst.Consultas) consultas ($($inst.Carga) busquedas)" "INFO"
                try { (Get-Process -Id $inst.Fijos.ProcessId -ErrorAction SilentlyContinue).Kill() } catch { }
                Start-Sleep -Seconds 2
                try {
                    $ctxN = Start-ACInstancia -Password $Password -BaseDatos $BaseDatos -Silencioso
                    $fj = Get-ACControlesFijos -ProcessId $ctxN.ProcessId
                    if (-not $fj) { throw "sin panel de criterios" }
                    $inst.Fijos = $fj
                    $inst.Consultas = 0
                    $inst.Carga = 0
                } catch {
                    Write-Paso "Instancia $($inst.Indice): no se pudo reiniciar ($($_.Exception.Message))" "WARN"
                    $inst.Estado = 'muerta'
                    continue
                }
            }

            # ---------------- libre: tomar el siguiente numero ----------------
            if ($inst.Estado -eq 'libre') {
                if ($cola.Count -gt 0) {
                    $num = $cola.Dequeue()
                    $inst.Numero   = $num
                    $inst.T0       = Get-Date
                    $inst.Intentos = 0
                    $inst.UltClic  = [datetime]::MinValue
                    try {
                        # Saca la instancia de un eventual modo menu (una pulsacion
                        # de ALT de otro proceso puede activarle la barra de menu,
                        # y ahi deja de atender los clics enviados por mensaje).
                        [void][W32]::PostMessage($inst.Fijos.Principal.Handle, 0x001F, [IntPtr]::Zero, [IntPtr]::Zero)

                        # activar + cebar con el radio MIN/MSISDN + escribir + BUSCAR
                        if (-not (Set-WindowFocus -Handle $inst.Fijos.Principal.Handle)) {
                            throw "no se pudo poner la ventana de AC en primer plano"
                        }
                        if ($inst.Fijos.Radio) {
                            Invoke-PostClick -Handle $inst.Fijos.Radio.Handle
                            Start-Sleep -Milliseconds 120
                        }
                        $leido = Set-CtrlText -Handle $inst.Fijos.Edit.Handle -Text $num
                        if ($leido -ne $num) { throw "el campo Criterio quedo en '$leido'" }
                        Invoke-PostClick -Handle $inst.Fijos.Buscar.Handle
                        # ver Invoke-ACClicActivado: el clic debe procesarse
                        # mientras esta instancia sigue en primer plano
                        Start-Sleep -Milliseconds 400
                        $inst.UltBuscar = Get-Date
                        $inst.Reenvios  = 0
                        $inst.Carga     = $inst.Carga + 1
                        $inst.Estado = 'esperando'
                    } catch {
                        Publicar -Inst $inst -Res ([pscustomobject]@{
                            Numero = $num; Encontrado = $false
                            Mensaje = "No se pudo lanzar la busqueda: $($_.Exception.Message)"; Filas = @()
                        })
                    }
                }
            }

            # ------------- esperando: que aparezca la ventana -----------------
            elseif ($inst.Estado -eq 'esperando') {
                $transcurrido = ((Get-Date) - $inst.T0).TotalSeconds
                if (Get-VentanaResultados -Fijos $inst.Fijos) {
                    $inst.Estado     = 'leyendo'
                    $inst.T0Leyendo  = Get-Date
                    $inst.Ctrls      = $null
                }
                else {
                    $resuelto = $false
                    # Los dialogos se revisan recien despues de unos segundos:
                    # enumerar ventanas en cada sondeo costaria mas que esperar.
                    if ($transcurrido -gt 3) {
                        if ((Get-ACDialogs -ProcessId $inst.Fijos.ProcessId).Count -gt 0) {
                            $msgs = Close-ACDialogs -ProcessId $inst.Fijos.ProcessId
                            Publicar -Inst $inst -Res ([pscustomobject]@{
                                Numero = $inst.Numero; Encontrado = $false
                                Mensaje = ($msgs -join ' | '); Filas = @()
                            })
                            $resuelto = $true
                        }
                    }

                    # Reenviar BUSCAR: con varias instancias el clic se pierde
                    # si otra roba el primer plano antes de que AC lo procese.
                    # Sin esto el numero terminaba como falso "NO ENCONTRADO".
                    if (-not $resuelto -and
                        ((Get-Date) - $inst.UltBuscar).TotalSeconds -gt 7 -and
                        $inst.Reenvios -lt $MaxReenvios) {

                        $inst.Reenvios++
                        if ($Trace) { Write-Host ("    [traza inst {0}] reenviando BUSCAR (intento {1})" -f $inst.Indice, $inst.Reenvios) -ForegroundColor DarkGray }
                        [void](Invoke-ACClicActivado -Fijos $inst.Fijos -Control $inst.Fijos.Buscar.Handle -Cebo $inst.Fijos.Radio.Handle)
                        $inst.UltBuscar = Get-Date
                        $inst.Carga     = $inst.Carga + 1
                    }

                    if (-not $resuelto -and $transcurrido -gt $TimeoutPorNumeroSeg) {
                        Publicar -Inst $inst -Res ([pscustomobject]@{
                            Numero = $inst.Numero; Encontrado = $false
                            Mensaje = "AC no respondio tras $($inst.Reenvios) reenvios en $TimeoutPorNumeroSeg s."
                            Filas = @()
                        })
                    }
                }
            }

            # ---------- leyendo: seleccionar la linea y leer la grilla --------
            elseif ($inst.Estado -eq 'leyendo') {
                $transcurrido = ((Get-Date) - $inst.T0).TotalSeconds
                $vent = Get-VentanaResultados -Fijos $inst.Fijos

                if (-not $vent) {
                    if ($transcurrido -gt $TimeoutPorNumeroSeg) {
                        Publicar -Inst $inst -Res ([pscustomobject]@{
                            Numero = $inst.Numero; Encontrado = $false
                            Mensaje = "La ventana de resultados desaparecio."; Filas = @()
                        })
                    }
                }
                else {
                    # Los handles del arbol y la grilla se resuelven una sola vez
                    # por consulta: enumerar los hijos cuatro veces por segundo
                    # durante toda la espera recargaba a AC sin necesidad.
                    if (-not $inst.Ctrls -or $inst.Ctrls.Vent -ne $vent.Handle) {
                        $k = Get-ChildHandles -Parent $vent.Handle
                        $inst.Ctrls = [pscustomobject]@{
                            Vent = $vent.Handle
                            Tv   = ($k | Where-Object { $_.Class -like 'TreeView*' } | Select-Object -First 1)
                            Lv   = ($k | Where-Object { $_.Class -like 'ListView*' } | Select-Object -First 1)
                        }
                    }
                    $tv = $inst.Ctrls.Tv
                    $lv = $inst.Ctrls.Lv

                    if ($tv -and $lv) {
                        # sondeo barato: un solo mensaje, sin memoria remota
                        $filas = Get-ListViewRowCount -Hwnd $lv.Handle

                        # -1 = la ventana no contesto dentro del plazo. Se le dan
                        # unos cuantos ciclos de gracia (AC se queda mudo un
                        # momento mientras Oracle le responde) pero no mas: una
                        # instancia colgada frena a todas las demas, y no vale la
                        # pena esperarla cuando reiniciarla cuesta ~20 s.
                        if ($filas -eq -1) {
                            $inst.Mudo = $inst.Mudo + 1
                            if ($inst.Mudo -ge $MaxSondeosMudos) {
                                Write-Paso ("Instancia $($inst.Indice): no responde hace {0} sondeos, se reinicia" -f $inst.Mudo) "WARN"
                                Publicar -Inst $inst -Res ([pscustomobject]@{
                                    Numero = $inst.Numero; Encontrado = $false
                                    Mensaje = "La instancia de AC dejo de responder."; Filas = @()
                                })
                                $inst.Mudo = 0
                                # forzar el reciclado en la proxima vuelta
                                $inst.Carga = [Math]::Max($inst.Carga, $ReciclarCada)
                            }
                            continue
                        }
                        $inst.Mudo = 0
                        if ($Trace) {
                            Write-Host ("    [traza inst {0}] t={1:N1}s filas={2} clics={3}" -f `
                                $inst.Indice, $transcurrido, $filas, $inst.Intentos) -ForegroundColor DarkGray
                        }

                        if ($filas -gt 0) {
                            $d = Get-ListViewData -Hwnd $lv.Handle
                            Publicar -Inst $inst -Res ([pscustomobject]@{
                                Numero = $inst.Numero; Encontrado = $true; Mensaje = ''
                                Columnas = $d.Columns; Filas = $d.Rows
                            })
                        }
                        else {
                            # Seleccionar la linea en el arbol, pero solo cuando
                            # el nodo ya existe: si se hace antes, AC ignora el
                            # clic y se pierde mas de un segundo por consulta.
                            $nodos = Get-TreeViewCount -Hwnd $tv.Handle
                            if ($nodos -eq -1) { $inst.Mudo = $inst.Mudo + 1; continue }
                            if (((Get-Date) - $inst.UltClic).TotalMilliseconds -gt 800) {
                                if ($nodos -gt 0) {
                                    [void](Invoke-ACClicActivado -Fijos $inst.Fijos -Control $tv.Handle -X 60 -Y 10)
                                    $inst.UltClic = Get-Date
                                    $inst.Intentos++
                                }
                            }

                            # Arbol vacio = la busqueda no trajo nada. Medido en
                            # 1409 consultas que si encontraron: mediana 10.2 s,
                            # p90 14 s, maximo 47 s. Seguir esperando 75 s por
                            # cada numero inexistente costo 87 de los 350 min de
                            # esa corrida y desgasta las instancias hasta
                            # tumbarlas, asi que se corta antes. Los cortados por
                            # aqui se reintentan luego con el plazo completo.
                            $esperaVacio = ((Get-Date) - $inst.T0Leyendo).TotalSeconds
                            if ($nodos -eq 0 -and $TimeoutVacioSeg -gt 0 -and $esperaVacio -gt $TimeoutVacioSeg) {
                                $msgs = Close-ACDialogs -ProcessId $inst.Fijos.ProcessId
                                Publicar -Inst $inst -Res ([pscustomobject]@{
                                    Numero = $inst.Numero; Encontrado = $false
                                    Mensaje = if ($msgs) { $msgs -join ' | ' } else { "La busqueda no devolvio ninguna linea." }
                                    Filas = @()
                                })
                            }
                            elseif ($transcurrido -gt $TimeoutPorNumeroSeg) {
                                $msgs = Close-ACDialogs -ProcessId $inst.Fijos.ProcessId
                                Publicar -Inst $inst -Res ([pscustomobject]@{
                                    Numero = $inst.Numero; Encontrado = $false
                                    Mensaje = if ($msgs) { $msgs -join ' | ' } else { "La busqueda no devolvio ninguna linea." }
                                    Filas = @()
                                })
                            }
                        }
                    }
                }
            }
        }

        if ($Instancias | Where-Object { $_.Estado -eq 'esperando' -or $_.Estado -eq 'leyendo' }) {
            Start-Sleep -Milliseconds $IntervaloMs
        }
        if (($Instancias | Where-Object { $_.Estado -ne 'muerta' }).Count -eq 0) {
            # Caida general. Antes se lanzaba una excepcion, con lo que la
            # corrida entera se perdia aunque faltaran pocos numeros; ahora se
            # intenta levantar de nuevo y, si no se puede, se devuelve lo
            # alcanzado para que quien llame lo guarde igual.
            $revividas = 0
            if ($Password -and $rondasRevivir -lt $MaxRevivir) {
                $rondasRevivir++
                Write-Paso "Todas las instancias cayeron; intento $rondasRevivir de $MaxRevivir para levantarlas" "WARN"
                foreach ($inst in $Instancias) {
                    try {
                        $ctxN = Start-ACInstancia -Password $Password -BaseDatos $BaseDatos -Silencioso
                        $fj = Get-ACControlesFijos -ProcessId $ctxN.ProcessId
                        if (-not $fj) { throw "sin panel de criterios" }
                        $inst.Fijos     = $fj
                        $inst.Carga     = 0
                        $inst.Consultas = 0
                        $inst.Numero    = $null
                        $inst.Intentos  = 0
                        $inst.Estado    = 'libre'
                        $revividas++
                        Write-Paso "Instancia $($inst.Indice): levantada de nuevo (PID $($ctxN.ProcessId))" "OK"
                    } catch {
                        Write-Paso "Instancia $($inst.Indice): no se pudo levantar ($($_.Exception.Message))" "WARN"
                    }
                }
            }
            if ($revividas -eq 0) {
                $incompleto = $true
                Write-Paso ("Sin instancias de AC: se devuelven los {0} de {1} numeros ya resueltos." -f `
                            $resultados.Count, $Numeros.Count) "ERROR"
                break
            }
        }
    }

    $swGlobal.Stop()
    $vivas = @($Instancias | Where-Object { $_.Estado -ne 'muerta' }).Count
    return [pscustomobject]@{
        Resultados        = $resultados
        SegundosTotales   = [Math]::Round($swGlobal.Elapsed.TotalSeconds, 1)
        SegundosPorNumero = if ($resultados.Count) { [Math]::Round($swGlobal.Elapsed.TotalSeconds / $resultados.Count, 2) } else { 0 }
        LatenciaMedia     = if ($tiempos.Count) { [Math]::Round(($tiempos | Measure-Object -Average).Average, 2) } else { 0 }
        InstanciasVivas   = $vivas
        # Incompleto = AC se quedo sin instancias antes de terminar la cola.
        # Los que faltan quedan en Pendientes para reintentarlos aparte.
        Incompleto        = $incompleto
        Pendientes        = @($cola.ToArray())
    }
}


