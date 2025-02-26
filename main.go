package main

import (
	"fmt"
	"log"
	"strings"
	"sync"

	"github.com/tealeg/xlsx"
)

// Função para alterar o nome dos vendedores em uma folha
func changeVendorName(sheet *xlsx.Sheet, oldName, newName string, ch chan<- error) {
	// Iterar sobre as linhas da folha
	for _, row := range sheet.Rows {
		// Verificar e alterar o nome dos vendedores
		if len(row.Cells) > 0 && row.Cells[0].String() == oldName {
			row.Cells[0].SetValue(newName)
		}
	}
}

// Funcao para alterar o metodo de pagamento de uma folha
func changePaymentMethod(sheet *xlsx.Sheet, oldName, newName string, ch chan<- error) {
	// Iterar sobre as linhas da folha
	for _, row := range sheet.Rows {
		// Verificar e alterar o metodo de pagamento (considerando que a coluna é a sexta, indexada em zero)
		if len(row.Cells) > 5 && strings.TrimSpace(row.Cells[6].String()) == oldName {
			row.Cells[6].SetValue(newName)
		}
	}
}

// Funcao para alterar o nome do vendor de uma folha (planilha2)
func changeVendorName2(sheet *xlsx.Sheet, oldName, newName string, ch chan<- error) {
	// Iterar sobre as linhas da folha
	for _, row := range sheet.Rows {
		// Verificar e alterar o nome dos vendedores
		if len(row.Cells) > 0 && row.Cells[0].String() == oldName {
			row.Cells[0].SetValue(newName)
		}
	}
}

func changeStatus(sheet *xlsx.Sheet, oldName, newName string, ch chan<- error) {
	// Iterar sobre as linhas da folha
	for _, row := range sheet.Rows {
		// Verificar e alterar o metodo de pagamento (considerando que a coluna é a sexta, indexada em zero)
		if len(row.Cells) > 3 && strings.TrimSpace(row.Cells[4].String()) == oldName {
			row.Cells[4].SetValue(newName)
		}
	}
}

func changeBlank(sheet *xlsx.Sheet, oldName, newName string, ch chan<- error) {
	// Iterar sobre as linhas da folha
	for _, row := range sheet.Rows {
		// Verificar e alterar o metodo de pagamento (considerando que a coluna é a sexta, indexada em zero)
		if len(row.Cells) > 0 && strings.TrimSpace(row.Cells[1].String()) == "" {
			row.Cells[1].SetValue(newName)
		}
	}
}

func changeBrand(sheet *xlsx.Sheet, oldName, newName string, ch chan<- error) {
	// Iterar sobre as linhas da folha
	for _, row := range sheet.Rows {
		// Verificar e alterar o metodo de pagamento (considerando que a coluna é a sexta, indexada em zero)
		if len(row.Cells) > 0 && row.Cells[1].String() == oldName {
			row.Cells[1].SetValue(newName)
		}
	}
}

func main() {
	// Ler a planilha do Excel
	file, err := xlsx.OpenFile("C:/Users/paulo/OneDrive/AnalisesMatCont/VendasJaneiro.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	file2, err := xlsx.OpenFile("C:/Users/paulo/OneDrive/AnalisesMatCont/estatisticaJaneiro.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	file3, err := xlsx.OpenFile("C:/Users/paulo/OneDrive/AnalisesMatCont/vendasPorMarcaJaneiro.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Criar um canal para erros
	errCh := make(chan error, len(file.Sheets))
	errCh2 := make(chan error, len(file2.Sheets))
	errCh3 := make(chan error, len(file3.Sheets))

	// Criar um grupo de espera para as goroutines
	var wg sync.WaitGroup
	var wg2 sync.WaitGroup
	var wg3 sync.WaitGroup

	// Iterar sobre as folhas da planilha
	for _, sheet := range file.Sheets {
		sheet.Rows[0].Cells[0].SetValue("Vendedor")
		// Incrementar o grupo de espera para cada folha
		wg.Add(3) // Incrementar para cada função (changeVendorName e changePaymentMethod)

		// Iniciar goroutines para alterar o nome dos vendedores e o método de pagamento na folha
		go func(sheet *xlsx.Sheet) {
			defer wg.Done()
			changeVendorName(sheet, "ALEX CHIARADIA                          -000070", "OUTROS", errCh)
			changeVendorName(sheet, "DANIEL TRABIJU                          -000555", "OUTROS", errCh)
			changeVendorName(sheet, "EDUARDO                                 -004415", "OUTROS", errCh)
			changeVendorName(sheet, "VINICIUS ZAINA                          -004673", "OUTROS", errCh)
			changeVendorName(sheet, "VANESSA ZAINA                           -005168", "OUTROS", errCh)
			changeVendorName(sheet, "PAULO ROBERTO CHIARADIA NETO            -006432", "OUTROS", errCh)
			changeVendorName(sheet, "SILVIA HELENA FERREIRA DO PRADO         -006621", "OUTROS", errCh)
			changeVendorName(sheet, "GLAUCIA GUEDES                          -004741", "OUTROS", errCh)
			changeVendorName(sheet, "ADENILZA MIRANDA                        -000017", "OUTROS", errCh)
			changeVendorName(sheet, "CAMILA TEIXEIRA                         -005208", "CAMILA", errCh)
			changeVendorName(sheet, "JOANA VITORIA MOREIRA DE JESUS          -006525", "JOANA", errCh)
			changeVendorName(sheet, "JULIA RIBEIRO BARBOSA                   -008639", "JULIA", errCh)
			changeVendorName(sheet, "JULIANA CHIARADIA                       -003820", "JULIANA", errCh)
			changeVendorName(sheet, "PATRICIA VASCONCELOS VITORIANO          -008737", "PATRICIA", errCh)
			changeVendorName(sheet, "PAULO AUGUSTO NOGUEIRA                  -008722", "PAULO", errCh)
			changeVendorName(sheet, "PEDRO HENRIQUE DE SOUZA DRAIHER         -006057", "PEDRO", errCh)
			changeVendorName(sheet, "TIAGO SILVA                             -006301", "TIAGO", errCh)
			changeVendorName(sheet, "VANESSA PRADO                           -004923", "VANESSA", errCh)
			changeVendorName(sheet, "WELLISON RODRIGUES                      -006572", "WELLISON", errCh)
			changeVendorName(sheet, "RAFAELA NEVES RAYMUNDO                  -008589", "RAFAELA", errCh)
			changeVendorName(sheet, "RAFAELA NEVES                           -008589", "RAFAELA", errCh)
			changeVendorName(sheet, "WAGNER FRANCA LOPES SAMPAIO             -006503", "WAGNER", errCh)
		}(sheet)

		go func(sheet *xlsx.Sheet) {
			defer wg.Done()
			changePaymentMethod(sheet, "01-PRAZO", "PRAZO", errCh)
			changePaymentMethod(sheet, "12-PRAZO PROMO", "PRAZO PROMO", errCh)
			changePaymentMethod(sheet, "06-EMPRESA CNPJ", "EMPRESA", errCh)
			changePaymentMethod(sheet, "03-PRO DEB/DIN", "A VISTA", errCh)
			changePaymentMethod(sheet, "02-PRO DEB/CRE", "A VISTA", errCh)
			changePaymentMethod(sheet, "02-A VISTA", "A VISTA", errCh)
			changePaymentMethod(sheet, "03-DIN/PIX RET", "A VISTA", errCh)
			changePaymentMethod(sheet, "05-PROMIS 1X", "PROMISSORIA", errCh)
			changePaymentMethod(sheet, "05-PROMIS 2X", "PROMISSORIA", errCh)
			changePaymentMethod(sheet, "05-PROMIS 3X", "PROMISSORIA", errCh)
			changePaymentMethod(sheet, "04-PROMO CRD 1X", "CARTAO", errCh)
			changePaymentMethod(sheet, "07-01X CAR CRED", "CARTAO", errCh)
			changePaymentMethod(sheet, "07-02X CAR CRED", "CARTAO", errCh)
			changePaymentMethod(sheet, "07-03X CAR CRED", "CARTAO", errCh)
			changePaymentMethod(sheet, "07-04X CAR CRED", "CARTAO", errCh)
			changePaymentMethod(sheet, "07-05X CAR CRED", "CARTAO", errCh)
			changePaymentMethod(sheet, "07-06X CAR CRED", "CARTAO", errCh)
			changePaymentMethod(sheet, "07-07X CAR CRED", "CARTAO", errCh)
			changePaymentMethod(sheet, "07-08X CAR CRED", "CARTAO", errCh)
			changePaymentMethod(sheet, "07-09X CAR CRED", "CARTAO", errCh)
			changePaymentMethod(sheet, "07-10X CAR CRED", "CARTAO", errCh)
			changePaymentMethod(sheet, "50-PROMO 12X", "CARTAO", errCh)
			changePaymentMethod(sheet, "08-CHQ VISTA", "CHEQUE", errCh)
			changePaymentMethod(sheet, "09-CHQ 0+1", "CHEQUE", errCh)
			changePaymentMethod(sheet, "09-CHQ 0+2", "CHEQUE", errCh)
			changePaymentMethod(sheet, "09-CHQ 0+3", "CHEQUE", errCh)
			changePaymentMethod(sheet, "10-CHQ 1+2", "CHEQUE", errCh)
		}(sheet)

		go func(sheet *xlsx.Sheet) {
			defer wg.Done()
			changeStatus(sheet, "Entregar", "ENTREGA", errCh)
			changeStatus(sheet, "Entregue", "ENTREGA", errCh)
			changeStatus(sheet, "Retirado", "RETIRA", errCh)
			changeStatus(sheet, "Pedido", "RETIRA", errCh)
		}(sheet)
	}

	// Aguardar a conclusão de todas as goroutines
	wg.Wait()

	// Fechar o canal de erros
	close(errCh)

	// Verificar se houve algum erro durante a execução das goroutines
	for err := range errCh {
		if err != nil {
			log.Fatal(err)
		}
	}

	// Salvar as alterações de volta no Excel
	err = file.Save("C:/Users/paulo/OneDrive/AnalisesMatCont/VendasJaneiroNormatizado.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Iterar sobre as folhas da planilha2
	for _, sheet2 := range file2.Sheets {

		// Incrementar o grupo de espera para cada folha
		wg2.Add(1) // Incrementar para cada função (changeVendorName e changePaymentMethod)

		// Iniciar goroutines para alterar o nome dos vendedores e o método de pagamento na folha
		go func(sheet2 *xlsx.Sheet) {
			defer wg2.Done()
			changeVendorName2(sheet2, "000070-ALEX CHIARADIA", "OUTROS", errCh2)
			changeVendorName2(sheet2, "000555-DANIEL TRABIJU", "OUTROS", errCh2)
			changeVendorName2(sheet2, "004415-EDUARDO", "OUTROS", errCh2)
			changeVendorName2(sheet2, "004673-VINICIUS ZAINA", "OUTROS", errCh2)
			changeVendorName2(sheet2, "005168-VANESSA ZAINA", "OUTROS", errCh2)
			changeVendorName2(sheet2, "006432-PAULO ROBERTO CHIA", "OUTROS", errCh2)
			changeVendorName2(sheet2, "006621-SILVIA HELENA FERR", "OUTROS", errCh2)
			changeVendorName2(sheet2, "004741-GLAUCIA GUEDES", "OUTROS", errCh2)
			changeVendorName2(sheet2, "000017-ADENILZA MIRANDA", "OUTROS", errCh2)
			changeVendorName2(sheet2, "005208-CAMILA TEIXEIRA", "CAMILA", errCh2)
			changeVendorName2(sheet2, "006525-JOANA VITORIA MORE", "JOANA", errCh2)
			changeVendorName2(sheet2, "008639-JULIA RIBEIRO BARB", "JULIA", errCh2)
			changeVendorName2(sheet2, "003820-JULIANA CHIARADIA", "JULIANA", errCh2)
			changeVendorName2(sheet2, "008737-PATRICIA VASCONCEL", "PATRICIA", errCh2)
			changeVendorName2(sheet2, "008722-PAULO AUGUSTO NOGU", "PAULO", errCh2)
			changeVendorName2(sheet2, "006057-PEDRO HENRIQUE DE", "PEDRO", errCh2)
			changeVendorName2(sheet2, "006301-TIAGO SILVA", "TIAGO", errCh2)
			changeVendorName2(sheet2, "004923-VANESSA PRADO", "VANESSA", errCh2)
			changeVendorName2(sheet2, "006572-WELLISON RODRIGUES", "WELLISON", errCh2)
			changeVendorName2(sheet2, "008589-RAFAELA NEVES", "RAFAELA", errCh2)
			changeVendorName2(sheet2, "006503-WAGNER FRANCA LOPE", "WAGNER", errCh2)
		}(sheet2)
	}
	// Aguardar a conclusão de todas as goroutines
	wg2.Wait()

	// Fechar o canal de erros
	close(errCh2)

	// Verificar se houve algum erro durante a execução das goroutines
	for err := range errCh2 {
		if err != nil {
			log.Fatal(err)
		}
	}
	// Salvar as alterações de volta no Excel
	err = file2.Save("C:/Users/paulo/OneDrive/AnalisesMatCont/estatisticaNormatizado.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Iterar sobre as folhas da planilha3
	for _, sheet3 := range file3.Sheets {
		sheet3.Rows[0].Cells[0].SetValue("Vendedor")
		// Incrementar o grupo de espera para cada folha
		wg3.Add(1) // Incrementar para cada função (changeVendorName e changePaymentMethod)

		// Iniciar goroutines para alterar o nome dos vendedores e o método de pagamento na folha
		go func(sheet3 *xlsx.Sheet) {
			defer wg3.Done()
			changeVendorName2(sheet3, "000070-ALEX CHIARADIA", "OUTROS", errCh3)
			changeVendorName2(sheet3, "000555-DANIEL TRABIJU", "OUTROS", errCh3)
			changeVendorName2(sheet3, "004415-EDUARDO", "OUTROS", errCh3)
			changeVendorName2(sheet3, "004673-VINICIUS ZAINA", "OUTROS", errCh3)
			changeVendorName2(sheet3, "005168-VANESSA ZAINA", "OUTROS", errCh3)
			changeVendorName2(sheet3, "006432-PAULO ROBERTO CHIARADIA NETO", "OUTROS", errCh3)
			changeVendorName2(sheet3, "006621-SILVIA HELENA FERREIRA DO PRADO", "OUTROS", errCh3)
			changeVendorName2(sheet3, "004741-GLAUCIA GUEDES", "OUTROS", errCh3)
			changeVendorName2(sheet3, "000017-ADENILZA MIRANDA", "OUTROS", errCh3)
			changeVendorName2(sheet3, "005208-CAMILA TEIXEIRA", "CAMILA", errCh3)
			changeVendorName2(sheet3, "006525-JOANA VITORIA MOREIRA DE JESUS", "JOANA", errCh3)
			changeVendorName2(sheet3, "008639-JULIA RIBEIRO BARBOSA", "JULIA", errCh3)
			changeVendorName2(sheet3, "003820-JULIANA CHIARADIA", "JULIANA", errCh3)
			changeVendorName2(sheet3, "008737-PATRICIA VASCONCELOS VITORIANO", "PATRICIA", errCh3)
			changeVendorName2(sheet3, "008722-PAULO AUGUSTO NOGUEIRA", "PAULO", errCh3)
			changeVendorName2(sheet3, "006057-PEDRO HENRIQUE DE SOUZA DRAIHER", "PEDRO", errCh3)
			changeVendorName2(sheet3, "006301-TIAGO SILVA", "TIAGO", errCh3)
			changeVendorName2(sheet3, "004923-VANESSA PRADO", "VANESSA", errCh3)
			changeVendorName2(sheet3, "006572-WELLISON RODRIGUES", "WELLISON", errCh3)
			changeVendorName2(sheet3, "008589-RAFAELA NEVES", "RAFAELA", errCh3)
			changeBlank(sheet3, "", "*****", errCh3)
			changeBrand(sheet3, "INHAPIM CEDR MESCL", "INHAPIM", errCh3)
			changeBrand(sheet3, "INHAPIM GARAPEIRA", "INHAPIM", errCh3)
			changeBrand(sheet3, "INHAPIM EUCALIPTO", "INHAPIM", errCh3)
			changeBrand(sheet3, "UNIDADE", "TIJOLO", errCh3)
			changeBrand(sheet3, "EUCALIPTO", "INHAPIM", errCh3)
		}(sheet3)

		// Aguardar a conclusão de todas as goroutines
		wg3.Wait()

		// Fechar o canal de erros
		close(errCh3)

		// Verificar se houve algum erro durante a execução das goroutines
		for err := range errCh3 {
			if err != nil {
				log.Fatal(err)
			}
		}

		// Salvar as alterações de volta no Excel
		err = file3.Save("C:/Users/paulo/OneDrive/AnalisesMatCont/VendasPorMarcaNormatizado.xlsx")
		if err != nil {
			log.Fatal(err)
		}

		fmt.Println("Arquivo salvo com sucesso!")
		fmt.Println("Alterações concluídas com sucesso!")
	}
}
