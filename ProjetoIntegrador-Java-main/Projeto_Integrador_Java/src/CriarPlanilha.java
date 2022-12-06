import jxl.Workbook;
import jxl.format.Colour;
import jxl.write.Number;

import java.io.File;
import java.util.ArrayList;

import jxl.write.DateFormat;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class CriarPlanilha {

    private String enderecoExcel = "H:/FPOO/teste/livea.xls";

    public CriarPlanilha(String excel){
        this.enderecoExcel = excel;
    }

    public void gerarPlanilha(){
        //System.out.println("Inicio");

        ArrayList<Object> lista = new ArrayList<Object>();

        Usuario usuario1 = new Usuario(
                "Café tipo 1",
                "R$ --",
                "Café preparado com o pó de café comercial\n"
        );
        Usuario usuario2 = new Usuario(
                "Tipo de grão 1\n",
                "R$ --",
                "Grão de café Arabica\n"
        );
        Usuario usuario3 = new Usuario(
                "Tipo de boneco 1",
                "R$ --",
                "Boneco anime chibi"
        );
        Usuario usuario4 = new Usuario(
                "Tipo de receita 1",
                "R$ --",
                "Receita 1 de café"
        );
        Usuario usuario5 = new Usuario(
                "Tipo de caneca 1",
                "R$ --",
                "Caneca de café sem decoração"
        );

        lista.add(usuario1);
        lista.add(usuario2);
        lista.add(usuario3);
        lista.add(usuario4);
        lista.add(usuario5);

        try {
            WritableWorkbook planilha = Workbook.createWorkbook(new File(
                    getEnderecoExcel()));
            // Adicionando o nome da aba
            WritableSheet aba = planilha.createSheet("ListaAlunos", 0);

            // Cabeçalhos
            String cabecalho[] = new String[5];
            cabecalho[0] = "ID";
            cabecalho[1] = "PRODUTO";
            cabecalho[2] = "VALOR";
            cabecalho[3] = "CATEGORIA";
            cabecalho[4] = "DATA DO CADASTRO";

            // Cor de fundo das celular
            Colour bckcolor = Colour.DARK_GREEN;
            WritableCellFormat cellFormat = new WritableCellFormat();
            cellFormat.setBackground(bckcolor);

            // Cor e tipo de fonte
            WritableFont fonte = new WritableFont(WritableFont.ARIAL);
            fonte.setColour(Colour.GOLD);
            cellFormat.setFont(fonte);

            // Write the Header to the Excel file
            for (int i = 0; i < cabecalho.length; i++) {
                Label label = new Label(i, 0, cabecalho[i]);
                aba.addCell(label);
                WritableCell cell = aba.getWritableCell(i, 0);
                cell.setCellFormat(cellFormat);
            }

            for (int linha = 0; linha < lista.size(); linha++) { // Número da linha

                Usuario usuario = (Usuario) lista.get(linha);

                jxl.write.Number number = new Number(0, linha, Double.parseDouble(usuario.getId()));

                aba.addCell(number);

                Label label = new Label(1, linha, usuario.getProduto());
                aba.addCell(label);

                label = new Label(2, linha, usuario.getValor());
                aba.addCell(label);

                label = new Label(3, linha, usuario.getCategoria());
                aba.addCell(label);

                DateFormat customDateFormat = new DateFormat(
                        "dd MMM yyyy hh:mm:ss");
                WritableCellFormat dateFormat = new WritableCellFormat(
                        customDateFormat);
                DateTime dateCell = new DateTime(4, linha, usuario.getDataCadastro(), dateFormat);

                aba.addCell(dateCell);
            }

            planilha.write();
            // Fecha o arquivo
            planilha.close();

        }catch(Exception e) {
            e.printStackTrace();
        }
        System.out.println("Fim");
    }

    public String getEnderecoExcel() {
        return enderecoExcel;
    }
}

