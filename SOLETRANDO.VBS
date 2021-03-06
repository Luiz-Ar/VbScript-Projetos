Option Explicit
dim jogador, nivel, acertos, points, palavra, n, audio, p_sorteio, p_escolhidas(15), cont_p, exist, repetir, jump
dim resp, aux

call carregar_jogo
sub carregar_jogo()
	set audio = createobject("SAPI.SPVOICE")
	audio.volume = 100
	nivel = "A"
	jump = 3
	acertos = 0
	cont_p = 0
	points = 0
	call nome_jogador
end sub

sub nome_jogador()
	jogador = InputBox("Nome do Jogador:", "SOLETRANDO")
	if jogador = False then
		wscript.quit
	elseif jogador = "" then
		call nome_jogador
	else	
		call inicio
	end if
end sub

sub inicio
	repetir = false
	if points >= 0 and points < 10000 then
		call nivelA
	elseif points >= 10000 and points < 100000  then
		call nivelB
	elseif points >= 100000 and points < 600000 then
		call nivelC
	end if
end sub

sub soletrar
	audio.rate = 1
	palavra = Inputbox("Nome do jogador: " & UCase(jogador) & "" & +vbNewLine & _
			"N�vel: " & nivel & "" & + vbNewLine & _
			"Qtde de acertos: " & acertos & "" & + vbNewLine & _
			"Pontua��o: " & points & "" & + vbNewLine & _
			"[O]uvir palavra novamente" & + vbNewLine & _
			"[P]ular palavra" & + vbNewLine & _
			"[E]ncerrar jogo", "SOLETRANDO " & Date &"")
		if palavra = False then
			wscript.quit
		end if
		palavra = LCase(palavra)
		if palavra = "O" or palavra = "o" then
			if repetir = false then
				repetir = true
				audio.rate = -4
				call ouvir
			elseif repetir = true then
				msgbox("Voc� j� repetiu essa palavra uma vez!")
			end if
			call soletrar
		elseif palavra = "P" or palavra = "p" then
			call pular
		elseif palavra = "E" or palavra = "e" then
			wscript.quit
		elseif palavra = "" then
			call soletrar
		else
			if palavra <> p_sorteio then
				msgbox("VOC� ERROU!! A palavra � " & UCase(p_sorteio) & ""), vbInformation + vbOKOnly, "ATEN��O"
				call fim
			elseif palavra = p_sorteio then
				msgbox("VOCE ACERTOU")
				acertos = acertos + 1
				if nivel = "A" then
					points = points + 1000
					if points = 5000 then
						msgbox("Voc� passou para o N�vel B e ganhou 5000 pontos!!!"), vbokonly + vbExclamation, "ATEN��O"
						nivel = "B"
						points = points + 5000
						call inicio
					end if
				elseif nivel = "B" then
					points = points + 10000
					if points = 60000 then
						msgbox("Voc� passou para o N�vel C e ganhou 40000 pontos!!!"), vbokonly + vbExclamation, "ATEN��O"
						nivel = "C"
						points = points + 40000
						call inicio
					end if
				elseif nivel = "C" then
					points = points + 100000
					if points = 600000 then
						msgbox("PARAB�NS!! VOC� VENCEU O JOGO!!"), vbExclamation + vbokonly, "PARAB�NS"
						call fim
					end if
				end if
			end if
			call inicio
		end if
end sub

sub nivelA
	randomize(second(time))
	n = int(rnd * 10) + 1
	'msgbox("N�mero sorteado: " & n & "")
	for aux = 0 to 4 step 1
		if n = p_escolhidas(aux) then
			call nivelA
			'msgbox("Este n�mero j� est� no vetor na posi��o " & aux & "")
		end if
	next
	p_escolhidas(cont_p) = n
	'msgbox("Numero cadastrado na posi��o " & cont_p & "")
	cont_p = cont_p + 1
	call palavras
end sub

sub nivelB
	randomize(second(time))
	n = int(rnd * 10) + 10
	for aux = 5 to 9 step 1
		if n = p_escolhidas(aux) then
			call nivelB
		end if
	next
	p_escolhidas(cont_p) = n
	cont_p = cont_p + 1
	call palavras
end sub

sub nivelC
	randomize(second(time))
	n = int(rnd * 10) + 20
	for aux = 10 to 14 step 1
		if n = p_escolhidas(aux) then
			call nivelC
		end if
	next
	p_escolhidas(cont_p) = n
	cont_p = cont_p + 1
	call palavras
end sub

sub palavras
		select case n
			case 1
				p_sorteio = "casa"
			case 2
				p_sorteio = "fugir"
			case 3
				p_sorteio = "mala"
			case 4
				p_sorteio = "coisa"
			case 5
				p_sorteio = "blusa"
			case 6
				p_sorteio = "luta"
			case 7
				p_sorteio = "jogo"
			case 8
				p_sorteio = "rede"
			case 9
				p_sorteio = "hoje"
			case 10
				p_sorteio = "viver"
			case 11
				p_sorteio = "l�grima"
			case 12
				p_sorteio = "despedida"
			case 13
				p_sorteio = "bancada"
			case 14
				p_sorteio = "tribunal"
			case 15
				p_sorteio = "biblioteca"
			case 16
				p_sorteio = "bombeiro"
			case 17
				p_sorteio = "profissional"
			case 18
				p_sorteio = "escritor"
			case 19
				p_sorteio = "general"
			case 20
				p_sorteio = "m�sica"
			case 21
				p_sorteio = "ardiloso"
			case 22
				p_sorteio = "insolente"
			case 23
				p_sorteio = "pern�stico"
			case 24
				p_sorteio = "agn�stico"
			case 25
				p_sorteio = "empedernido"
			case 26
				p_sorteio = "iconoclasta"
			case 27
				p_sorteio = "mendacioso"
			case 28
				p_sorteio = "progn�stico"
			case 29
				p_sorteio = "vicissitude"
			case 30
				p_sorteio = "zoom�rfico"
	end select
	call ouvir
	call soletrar
end sub

sub ouvir
	audio.speak("A palavra � "& p_sorteio & "")
end sub

sub pular
	if jump > 0 then
		resp = msgbox("Deseja realmente pular a palavra?", vbYesNo + vbQuestion, "ATEN��O")
		if resp = vbyes then
			jump = jump - 1
			msgbox("Voc� pulou a palavra e s� pode pular mais " & jump & " vezes"), vbInformation + vbOKOnly, "ATEN��O"
			call inicio
		else
			call soletrar
		end if
	else
		msgbox("Voc� j� utilizou os 3 pulos!!"), vbExclamation + vbOKOnly, "ATEN��O"
		call soletrar
	end if
end sub

sub fim
	resp = msgbox("Deseja jogar novamente?", vbyesno + vbQuestion, "ATEN��O")
	if resp = vbyes then
		call carregar_jogo
	else
		wscript.quit
	end if
end sub