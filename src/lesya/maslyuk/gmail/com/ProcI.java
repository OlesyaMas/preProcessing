package lesya.maslyuk.gmail.com;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JRadioButton;
import java.awt.BorderLayout;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.JToolBar;
import javax.swing.JButton;
import javax.swing.JTextArea;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.awt.event.ActionEvent;
import javax.swing.JTextPane;
import javax.swing.UIManager;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.JCheckBox;
import javax.swing.JFileChooser;

import java.awt.SystemColor;
import javax.swing.JLabel;
import java.awt.Font;
import java.awt.Color;
import java.awt.Dimension;
import javax.swing.JToggleButton;
import javax.swing.JProgressBar;
import javax.swing.JTabbedPane;
import javax.swing.JMenuBar;
import javax.swing.JMenu;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.AbstractAction;
import javax.swing.Action;

public class ProcI {

	private JFrame frame;
	private JTextField textField;
	JFileChooser fileChooser = new JFileChooser();
	static String inputFileFullName = "D:\\study\\java_proj\\test\\Книга1_200_code.xlsx";
	static String outputFile = "D:\\study\\java_proj\\test\\Книга1_out_test.xlsx";
	
	///static String inputFileFullName = "D:\\study\\java_proj\\data\\xls\\Книга1_200.xls";
	///static String outputFile = "D:\\study\\java_proj\\data\\xls\\Книга1_200_out.xls";
	
	//private final Action action = new SwingAction();
	
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ProcI window = new ProcI();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public ProcI() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 857, 452);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		/*textField = new JTextField();
		textField.setBounds(256, 62, 338, 20);
		frame.getContentPane().add(textField);
		textField.setColumns(10);*/
		 
		 fileChooser.setDialogTitle("Виберіть файл");
		 frame.getContentPane().add(fileChooser);
		 fileChooser.setVisible(true);
		
		JButton btnNewButton = new JButton("\u0410\u043D\u0430\u043B\u0456\u0437\u0443\u0432\u0430\u0442\u0438");
		btnNewButton.addActionListener(new ProcessFileHandler());
		
		btnNewButton.setForeground(new Color(128, 0, 0));
		btnNewButton.setBackground(UIManager.getColor("ToolTip.background"));
		btnNewButton.setFont(new Font("Sitka Text", Font.BOLD, 14));
		btnNewButton.setBounds(74, 355, 499, 33);
		frame.getContentPane().add(btnNewButton);
		
		JCheckBox cbCoding = new JCheckBox("\u041A\u043E\u0434\u0443\u0432\u0430\u043D\u043D\u044F \u044F\u043A\u0456\u0441\u043D\u0438\u0445 \u0437\u043D\u0430\u0447\u0435\u043D\u044C");
		cbCoding.setFont(new Font("Tahoma", Font.BOLD, 11));
		cbCoding.addActionListener(new CodingHandler());
		cbCoding.setBounds(74, 316, 213, 23);
		frame.getContentPane().add(cbCoding);
		
		JCheckBox checkBox = new JCheckBox("\u0410\u043D\u0430\u043B\u0456\u0437 \u0434\u0443\u0431\u043B\u044E\u044E\u0447\u0438\u0445 \u0440\u044F\u0434\u043A\u0456\u0432");
		checkBox.setFont(new Font("Trebuchet MS", Font.PLAIN, 12));
		checkBox.setBounds(489, 275, 195, 23);
		frame.getContentPane().add(checkBox);
		
		JCheckBox cbRemoveMinMax = new JCheckBox("\u0412\u0438\u0434\u0430\u043B\u0435\u043D\u043D\u044F \u0430\u0443\u0442\u043B\u0430\u0439\u043D\u0456\u0432 (Min/Max)");
		cbRemoveMinMax.addActionListener(new RemoveMinMaxHandler());
		cbRemoveMinMax.setFont(new Font("Tahoma", Font.BOLD, 11));
		cbRemoveMinMax.setBounds(74, 234, 260, 23);
		frame.getContentPane().add(cbRemoveMinMax);
		
		JCheckBox checkBox_3 = new JCheckBox("\u0410\u043D\u0430\u043B\u0456\u0437 \u043D\u0435\u0432\u0456\u0434\u043F\u043E\u0432\u0456\u0434\u043D\u043E\u0441\u0442\u0435\u0439 \u0444\u043E\u0440\u043C\u0430\u0442\u0443");
		checkBox_3.setFont(new Font("Trebuchet MS", Font.PLAIN, 12));
		checkBox_3.setBounds(489, 228, 235, 23);
		frame.getContentPane().add(checkBox_3);
		
		JCheckBox checkBox_4 = new JCheckBox("\u0410\u043D\u043E\u043D\u0456\u043C\u0456\u0437\u0430\u0446\u0456\u044F \u0434\u0430\u043D\u0438\u0445");
		checkBox_4.setFont(new Font("Trebuchet MS", Font.PLAIN, 12));
		checkBox_4.setBounds(489, 189, 195, 23);
		frame.getContentPane().add(checkBox_4);
		
		JLabel label = new JLabel("\u0412\u043A\u0430\u0436\u0456\u0442\u044C \u0448\u043B\u044F\u0445 \u0434\u043E \u0444\u0430\u0439\u043B\u0443:");
		label.setBackground(Color.ORANGE);
		label.setFont(new Font("Candara", Font.BOLD, 16));
		label.setBounds(69, 35, 172, 20);
		frame.getContentPane().add(label);
		
		JLabel label_2 = new JLabel("\u041E\u0431\u0435\u0440\u0456\u0442\u044C \u0432\u0438\u0434 \u043E\u0431\u0440\u043E\u0431\u043A\u0438:");
		label_2.setPreferredSize(new Dimension(124, 14));
		label_2.setMinimumSize(new Dimension(124, 14));
		label_2.setMaximumSize(new Dimension(124, 14));
		label_2.setFont(new Font("Candara", Font.BOLD, 16));
		label_2.setBackground(Color.ORANGE);
		label_2.setBounds(74, 158, 167, 14);
		frame.getContentPane().add(label_2);
		
		JButton button_1 = new JButton("\u0417\u0430\u0432\u0430\u043D\u0442\u0430\u0436\u0438\u0442\u0438 \u0444\u0430\u0439\u043B");
		button_1.addActionListener(new LoadFileHandler());
		
		button_1.setForeground(new Color(128, 0, 0));
		button_1.setFont(new Font("Sitka Text", Font.BOLD, 14));
		button_1.setBackground(SystemColor.info);
		button_1.setBounds(69, 107, 173, 33);
		frame.getContentPane().add(button_1);
		
		JButton button_2 = new JButton("\u0417\u0431\u0435\u0440\u0435\u0433\u0442\u0438");
		button_2.setForeground(new Color(128, 0, 0));
		button_2.setFont(new Font("Sitka Text", Font.BOLD, 14));
		button_2.setBackground(SystemColor.info);
		button_2.setBounds(616, 355, 173, 33);
		frame.getContentPane().add(button_2);
		
		JCheckBox cbAverage = new JCheckBox("\u0417\u0430\u043C\u0456\u043D\u0430 \u043F\u0440\u043E\u043F\u0443\u0449\u0435\u043D\u0438\u0445 \u0437\u043D\u0430\u0447\u0435\u043D\u044C \u043D\u0430 \u0441\u0435\u0440\u0435\u0434\u043D\u0454");
		cbAverage.setFont(new Font("Tahoma", Font.BOLD, 11));
		cbAverage.addActionListener(new CalcAverageHandler());
		//cbAverage.setAction(action);
		cbAverage.setBounds(74, 188, 282, 23);
		frame.getContentPane().add(cbAverage);
		
		JCheckBox cbRemoveEmptyCells = new JCheckBox("\u0412\u0438\u0434\u0430\u043B\u0435\u043D\u043D\u044F \u0440\u044F\u0434\u043A\u0456\u0432 \u0437 \u043F\u0443\u0441\u0442\u0438\u043C\u0438 \u044F\u0447\u0435\u0439\u043A\u0430\u043C\u0438");
		cbRemoveEmptyCells.setFont(new Font("Tahoma", Font.BOLD, 11));
		cbRemoveEmptyCells.addActionListener(new RemoveEmptyHandler());
		cbRemoveEmptyCells.setBounds(74, 274, 282, 23);
		frame.getContentPane().add(cbRemoveEmptyCells);
	}
	
	
	private final class CalcAverageHandler implements ActionListener {
		public void actionPerformed(ActionEvent e) {
			JCheckBox cbAverage = (JCheckBox)e.getSource();
			if(cbAverage.isSelected()){
				//JOptionPane.showMessageDialog(frame, "JCheckBox is selected");
				Processing.CALC_AVERAGE = true;
			}else{
				//JOptionPane.showMessageDialog(frame, "JCheckBox is NOT selected");
				Processing.CALC_AVERAGE = false;
			}
		}
	}

	private final class RemoveEmptyHandler implements ActionListener {
		public void actionPerformed(ActionEvent e) {
			JCheckBox cbRemoveEmptyCells = (JCheckBox)e.getSource();
			if(cbRemoveEmptyCells.isSelected()){
				Processing.DELETE_EMPTY = true;
			}else{
				Processing.DELETE_EMPTY = false;
			}
		}
	}
	
	private final class RemoveMinMaxHandler implements ActionListener {
		public void actionPerformed(ActionEvent e) {
			JCheckBox cbRemoveMinMax = (JCheckBox)e.getSource();
			if(cbRemoveMinMax.isSelected()){
				Processing.REMOVE_MIN_MAX = true;
			}else{
				Processing.REMOVE_MIN_MAX = false;
			}
		}
	}
	
	private final class CodingHandler implements ActionListener {
		public void actionPerformed(ActionEvent e) {
			JCheckBox cbCoding = (JCheckBox)e.getSource();
			if(cbCoding.isSelected()){
				Processing.CODING = true;
			}else{
				Processing.CODING = false;
			}
		}
	}
	

	private final class LoadFileHandler implements ActionListener {
		public void actionPerformed(ActionEvent e) {
			
			System.out.println("Load is pressed");
			int returnVal = fileChooser.showOpenDialog(frame);
			/*Docs importDocs = new Docs();
			importDocs.*/
		    //int returnVal = chooser.showOpenDialog(chooser);
			if (returnVal == JFileChooser.APPROVE_OPTION) {
				File inputFile = fileChooser.getSelectedFile();
				//TODO move to Docs read file + show message in label
				inputFileFullName = inputFile.getAbsolutePath().toString();
			}
		}
	}


	private final class ProcessFileHandler implements ActionListener {
		public void actionPerformed(ActionEvent e) {
			try {
				//Docs.readFile(inputFile);
				Processing.readFromExcel(inputFileFullName, outputFile);								
			} catch (FileNotFoundException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (IOException e2) {
				// TODO Auto-generated catch block
				e2.printStackTrace();
			} catch (Exception e3) {
				// TODO Auto-generated catch block
				e3.printStackTrace();
			}
			System.out.println("File completed !!!!!!!!!!!!!!!!!!!!!!!!!");
		}
	}
	private class SwingAction extends AbstractAction {
		public SwingAction() {
			putValue(NAME, "SwingAction");
			putValue(SHORT_DESCRIPTION, "Some short description");
		}
		public void actionPerformed(ActionEvent e) {
		}
	}
}
