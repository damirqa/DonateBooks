package damirqa.com.github.gui;

import java.awt.BorderLayout;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Toolkit;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import javax.swing.JLabel;
import javax.swing.JTextField;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Date;
import java.util.List;
import java.awt.event.ActionEvent;
import javax.swing.JTable;
import javax.swing.border.LineBorder;
import java.awt.Color;

@SuppressWarnings("serial")
public class Window extends JFrame {

	private int width = 600;
	private int height = 400;
	
	private JLabel fioLabel;
	private JTextField fioField;
	
	private JLabel addressLabel;
	private JTextField addressField;
	
	private JLabel passportLabel;
	private JTextField passportField;
	
	final private DefaultTableModel model;
	private JTable table;
	private JButton addRowsButton;
	private JButton removeRowsButton;
	
	private JButton formAkt;
	
	private JPanel contentPanel;
	
	

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Window frame = new Window();
					frame.setVisible(true);					
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public Window() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setTitle("Заявка");
		setWindowToCenter();
		
		contentPanel = new JPanel();
		contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPanel);
		contentPanel.setLayout(null);
		
		fioLabel = new JLabel("ФИО:");
		fioLabel.setFont(new Font("Segoe UI", Font.PLAIN, 16));
		fioLabel.setBounds(10, 11, 46, 14);
		contentPanel.add(fioLabel);
		
		fioField = new JTextField();
		fioField.setBounds(10, 36, 200, 20);
		fioField.setFont(new Font("Segoe UI", Font.PLAIN, 16));
		contentPanel.add(fioField);
		fioField.setColumns(10);
		
		addressLabel = new JLabel("Адрес регистрации");
		addressLabel.setFont(new Font("Segoe UI", Font.PLAIN, 16));
		addressLabel.setBounds(226, 8, 150, 20);
		contentPanel.add(addressLabel);
		
		addressField = new JTextField();
		addressField.setFont(new Font("Segoe UI", Font.PLAIN, 16));
		addressField.setBounds(220, 36, 200, 20);
		contentPanel.add(addressField);
		
		passportLabel = new JLabel("Паспортные данные");
		passportLabel.setFont(new Font("Segoe UI", Font.PLAIN, 16));
		passportLabel.setBounds(431, 8, 150, 20);
		contentPanel.add(passportLabel);
		
		passportField = new JTextField();
		passportField.setFont(new Font("Segoe UI", Font.PLAIN, 16));
		passportField.setBounds(430, 36, 200, 20);
		contentPanel.add(passportField);
				
		model = new DefaultTableModel();
		model.addColumn(new Object[]{"name, author"}, new Object[]{"Название, автор"});
		model.addColumn(new Object[]{"place, publish"}, new Object[]{"Место издания, изд-во"});
		model.addColumn(new Object[]{"year"}, new Object[]{"Год издания"});
		model.addColumn(new Object[]{"price"}, new Object[]{"Цена 1экз./руб."});
		model.addColumn(new Object[]{"col"}, new Object[]{"Кол-во экземпляров"});
		
		table = new JTable(model);
		table.setEnabled(true);
		table.setBorder(new LineBorder(new Color(0, 0, 0)));
		table.setBounds(10, 67, 620, 154);
		contentPanel.add(table);
		
		addRowsButton = new JButton("Добавить строчку");
		addRowsButton.setFont(new Font("Segoe UI", Font.PLAIN, 16));
		addRowsButton.setBounds(10, 230, 620, 25);
		contentPanel.add(addRowsButton);
		
		removeRowsButton = new JButton("Удалить строчку");
		removeRowsButton.setFont(new Font("Segoe UI", Font.PLAIN, 16));
		removeRowsButton.setBounds(10, 260, 620, 25);
		contentPanel.add(removeRowsButton);
		
		formAkt = new JButton("Сформировать акт");
		formAkt.setFont(new Font("Segoe UI", Font.PLAIN, 16));
		formAkt.setBounds(10, 290, 620, 25);
		contentPanel.add(formAkt);
		
		addRowsButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				model.addRow(new Object[model.getColumnCount()]);
			}
		});
		
		removeRowsButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				model.removeRow(model.getRowCount() - 1);
			}
		});
				
		formAkt.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					@SuppressWarnings("resource")
					XWPFDocument document = new XWPFDocument(OPCPackage.open(new File(getClass().getClassLoader().getResource("temp.docx").getPath())));
					XWPFTable table = document.getTables().get(0);
										
					int rows = model.getRowCount();
					int columns = model.getColumnCount();
					
					int total = 0;
					int price = 0;
					int amount = 0;
					
					for (int row = 1; row < rows; row++) {
						XWPFTableRow newRow = table.createRow();
						newRow.getCell(0).setText(row + "\n");
						for (int column = 1; column < columns + 1; column++) {
							newRow.getCell(column).setText(model.getValueAt(row, column - 1) + "\n");
							if (column == 4) price = Integer.parseInt((String) model.getValueAt(row, column - 1));
							if (column == 5) amount = Integer.parseInt((String) model.getValueAt(row, column - 1));
						}
						total += price * amount;
					}
					
					XWPFTableRow row = table.createRow();
					row.getCell(0).setText("Итого");
					row.getCell(5).setText(String.valueOf(total));
					
					setDate(document);
					
					updateData(document, "FIO", fioField);
					updateData(document, "Address", addressField);
					updateData(document, "Passport", passportField);
					updateTotal(document, "TOTAL", total);
					updateshortfio(document, "SHORT", fioField);
					
					//old path: C:\\Users\\kacer\\Desktop\\akt.docx
					FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.dir") + "akt1.docx");
					Path path = Paths.get("akt1.docx");
					document.write(outputStream);
					outputStream.close();
					
					System.out.println(System.getProperty("user.dir"));
					
					if (Desktop.isDesktopSupported()) {
						   Desktop.getDesktop().open(new File(System.getProperty("user.dir") + "akt1.docx"));
					}				
				} catch (InvalidFormatException e1) {
					e1.printStackTrace();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		});
		
	}
	
	private void setDate(XWPFDocument document) {
		Date date = new Date();
		String month = null;
				switch(date.getMonth()) {
					case 0: month = "Января";
						break;
					case 1: month = "Февраля";
						break;
					case 2: month = "Марта";
						break;
					case 3: month = "Апреля";
						break;
					case 4: month = "Мая";
						break;
					case 5: month = "Июня";
						break;
					case 6: month = "Июля";
						break;
					case 7: month = "Августа";
						break;
					case 8: month = "Сентебря";
						break;
					case 9: month = "Октября";
						break;
					case 10: month = "Ноября";
						break;
					case 11: month = "Декабря";
					break;
					
				}
		String setDate = "«"+date.getDate() + "» " + month + " " + (date.getYear() + 1900);
		
		for (XWPFParagraph p : document.getParagraphs()) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            if (text != null && text.contains("date")) {
		                text = text.replace("date", setDate);
		                r.setText(text, 0);
		            }
		        }
		    }
		}		
	}
	
	private void updateData(XWPFDocument document, String find_text, JTextField field) {
		for (XWPFParagraph p : document.getParagraphs()) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            if (text != null && text.contains(find_text)) {
		                text = text.replace(find_text, field.getText());
		                r.setText(text, 0);
		            }
		        }
		    }
		}
	}
	
	private void updateTotal(XWPFDocument document, String find_text, int total) {
		for (XWPFParagraph p : document.getParagraphs()) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            if (text != null && text.contains(find_text)) {
		                text = text.replace(find_text, total+"");
		                r.setText(text, 0);
		            }
		        }
		    }
		}
	}
	
	private void updateshortfio(XWPFDocument document, String find_text, JTextField field) {
		String fio = field.getText();
		String fioshort[] = fio.split(" ");
		String fios = fioshort[1].substring(0, 1) + ". " + fioshort[2].substring(0, 1) + ". " + fioshort[0];
		
		for (XWPFParagraph p : document.getParagraphs()) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            if (text != null && text.contains(find_text)) {
		                text = text.replace(find_text, fios);
		                r.setText(text, 0);
		            }
		        }
		    }
		}
	}
	
	private void setWindowToCenter() {
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		int x = (screenSize.width - this.width) / 2;
		int y = (screenSize.height - this.height) / 2;
		setBounds(x, y, 655, 358);
	}
}
