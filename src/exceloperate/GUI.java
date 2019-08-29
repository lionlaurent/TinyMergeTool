package exceloperate;

import java.awt.*;
import java.awt.event.*;
import java.util.ArrayList;
import java.util.List;

import javax.swing.*;


public class GUI {
	
	public static void main(String[] args) {
		
		EventQueue.invokeLater(() ->
		{
			MakeFrame frame = new MakeFrame();
			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
			frame.setVisible(true);
			frame.setTitle("IPD MergeTool@COMAC/jscbw/lionlaurent ver1.0");
			
		});
	}
	
	

}

class MakeFrame extends JFrame {
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	JPanel northPanel;
	JTextArea textArea1;
	JTextArea textArea2;
	private JFileChooser fileChooser = new JFileChooser();
	private static final int DEFAULT_WIDTH = 800;
	private static final int DEFAULT_HEIGHT = 640;
	
	List<String> lst = new ArrayList<String> ();
	
	public MakeFrame() 
	{	
		setSize(DEFAULT_WIDTH, DEFAULT_HEIGHT);
		
		northPanel = new JPanel();
		JButton b1 = new JButton("Choose Table"); 
		JButton b2 = new JButton("Input"); 
		northPanel.setLayout(new GridLayout(1, 2));
		northPanel.add(b1);
		northPanel.add(b2);
		add(northPanel, BorderLayout.NORTH);
		
		textArea1 = new JTextArea(8, 20);
		JScrollPane scrollPane1 = new JScrollPane(textArea1);
		textArea2 = new JTextArea(8, 20);
		JScrollPane scrollPane2 = new JScrollPane(textArea2);
		JPanel centerPanel = new JPanel();
		centerPanel.setLayout(new GridLayout(1, 2));
		centerPanel.add(scrollPane1);
		centerPanel.add(scrollPane2);
	
		add(centerPanel, BorderLayout.CENTER);
		
		
		JPanel southPanel = new JPanel();
		JButton b3 = new JButton("MakeOutput"); 
		southPanel.add(b3);
		add(southPanel, BorderLayout.SOUTH);
		
		ColorAction yellowaction = new ColorAction(Color.YELLOW);
		b1.addActionListener(new ChooseAction());
		b2.addActionListener(new InputAction());
		b3.addActionListener(new OutputAction());
	}
	
	private class ColorAction implements ActionListener {
		
		private Color backgroundColor;
		
		public ColorAction (Color c) {
			backgroundColor = c;
		}
		
		public void actionPerformed (ActionEvent event) {
			northPanel.setBackground(backgroundColor);
		}
	}
	
    
	private class ChooseAction implements ActionListener {
		
		public void actionPerformed (ActionEvent event) {
			int i = fileChooser.showOpenDialog(getContentPane());
			if (i == JFileChooser.APPROVE_OPTION) {
				textArea1.setText(fileChooser.getSelectedFile().getAbsolutePath());
				
			}
		}
	}
	
	private class InputAction implements ActionListener {
    	
		public void actionPerformed (ActionEvent event) {
			textArea2.append(textArea1.getText() + "\n");
			lst.add(textArea1.getText());
			textArea1.setText("");
			
		}
	}
    
    private class OutputAction extends Test5 implements ActionListener {
    	
		public void actionPerformed (ActionEvent event) {
			String savepath = null;
			int i = fileChooser.showSaveDialog(getContentPane());
			if (i == JFileChooser.APPROVE_OPTION) {
				savepath = fileChooser.getSelectedFile().getAbsolutePath();
				//System.out.println(savepath);
				String[] path = new String[lst.size()];
				for (int j = 0; j < path.length; j++) {
					path[j] = lst.get(j);
				}
				try {
					makeOutput(path, savepath);
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				lst.clear();
				textArea2.setText("");
			}
		}
	}
	
}
