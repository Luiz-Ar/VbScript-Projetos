Option Explicit
dim numero(10), escolha, audio, n, cont, aux, resp, cont2, nivel, foiOuvido, msg, player
dim jogando

jogando = 0

call carregar_audio

sub carregar_audio()
	set audio = createobject("SAPI.SPVOICE")
	audio.volume = 100
	audio.rate = 1
	call reset
end sub

sub reset
	nivel = "A"
	cont = 0
	cont2 = 0
	for aux = 0 to 9 step 1
		numero(aux) = Empty
	next
	if jogando = 0 then
		call jogador
	else
		call inicio
	end if
end sub

sub jogador()
	player = inputbox("Nome do Jogador", "JOGADOR")
	if player = False then
		wscript.quit
	elseif player = "" then
		msgbox("Digite o nome do jogador!"), 0 + 48, "ATENÇÃO"
		call jogador
	else
		call inicio
	end if
end sub

sub inicio()
	foiOuvido = False
	msg = False
	randomize(second(time))
	n = int(rnd * 100) + 1
	n = CStr(n)
	numero(cont) = n
	cont = cont + 1
	if cont = 11 then
		msgbox("PARABÉNS!!!" & vbNewLine & "Você venceu o jogo!!!"), 0 + 48, "ATENÇÃO"
		call fim
	end if
	select case cont
		case 3
			nivel = "A"
		case 4
			nivel = "B"
		case 5
			nivel = "C"
		case 6
			nivel = "D"
		case 7
			nivel = "E"
		case 8
			nivel = "F"
		case 9
			nivel = "G"
		case 10
			nivel = "H"
	end select
	if cont >= 3 then
		msgbox("NÍVEL " & nivel & ""), 0, "NÍVEL " & nivel &""
		call comparar
	end if
	call inicio
end sub

sub comparar() 
	call ouvir
	'wscript.sleep 3000
	escolha = inputbox(ucase("Jogador: " & player &"" & vbNewLine & vbNewLine & _
							 "Nível: " & nivel &"" & vbNewLine & vbNewLine & _
							 "Entre os números na sequencia:" & vbNewLine & vbNewLine & vbNewLine & _
					   "(SEPARADOS POR VÍRGULA E SEM ESPAÇOS)"), "ATENÇÃO")
	
	if escolha = False then
		wscript.quit
	elseif escolha <> "" then
		escolha = split(escolha, ",")
		if (UBound(escolha) + 1) > cont then
			msgbox("Você digitou número A MAIS!" & vbNewLine & vbNewLine & "Digite a quantidade exata de NÚMEROS!"), 0 + 64, "ATENÇÃO"
			call comparar
		end if
		cont2 = 0 
		do while cont2 < cont
			On Error Resume Next 
			if IsNumeric(escolha(cont2)) = False then
				msgbox("Você digitou número A MENOS ou não digitou números!" & vbNewLine & vbNewLine & _
					   "Digite a Sequencia de NÚMEROS!"), 0 + 64, "ATENÇÃO"
				call comparar
			end if
			cont2 = cont2 + 1
		loop
	else
		msgbox("PREENCHA O CAMPO COM A SEQUENCIA NUMÉRICA!!"), 0 + 64, "ATENÇÃO"
		call comparar
	end if
	cont2 = 0
	do while cont2 < cont
		if escolha(cont2) <> numero(cont2) then
			msgbox("O " & cont2 + 1 & "º número na sequencia está errado!" & vbNewLine & vbNewLine & _
				   "O número correto é " & numero(cont2) & " e você digitou " & escolha(cont2) & "!"), 0 + 16, "ATENÇÃO"
			msgbox("Você chegou ao Nível " & nivel & ""), 0 + 64, "ATENÇÃO"
			call fim
		end if
		cont2 = cont2 + 1
	loop
	if msg = False then
		msgbox("A SEQUENCIA ESTÁ CORRETA!!!"), 0 + 64, "ATENÇÃO"
		msg = True
	end if
end sub

sub ouvir
	if foiOuvido = False then
		for aux = 0 to cont step 1
			audio.speak(numero(aux))
		next
		foiOuvido = True
	end if
end sub

sub fim()
	resp = msgbox("Deseja jogar novamente?", 4 + 32, "ATENÇÃO")
	if resp = vbyes then
		jogando = 1
		call reset
	else
		wscript.quit
	end if
end sub