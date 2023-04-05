package com.kgd.swing;

import com.kgd.poi.Converter;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.ExecutionException;
import java.util.prefs.Preferences;
import java.util.stream.Collectors;

public class P2eConverter extends JFrame {

    private JTextField pptFilesText;
    private JButton pptSelectButton;
    private JLabel pptFilesLabel;
    private JButton generateButton;
    private JButton excelSelectButton;
    private JTextField excelFilesText;
    private JLabel excelFilesLabel;

    private File[] pptFiles;
    private File excelFileDir;

    private JPanel mainPanel;
    private JLabel resultLabel;
    private JTextArea resultTxt;

    public P2eConverter() {
        setContentPane(mainPanel);
        setTitle("PPT to Excel");
        setSize(600, 300);
        setResizable(false);
        setLocationRelativeTo(null);
        setVisible(true);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        resultTxt.setBackground(new Color(242, 242, 242));

        Preferences pref = Preferences.userRoot();

        String pptPath = pref.get("P2E_DEFAULT_PPT_PATH", System.getProperty("user.home"));
        String excelPath = pref.get("P2E_DEFAULT_EXCEL_PATH", System.getProperty("user.home"));

        pptSelectButton.addActionListener(e -> {
            JFileChooser chooser = new JFileChooser(new File(pptPath));
            chooser.setMultiSelectionEnabled(true);
            chooser.setDialogTitle("Select PPT files");
            chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            chooser.setFileFilter(new FileNameExtensionFilter("PowerPoint Files(*.pptx)", "pptx"));
            chooser.setAcceptAllFileFilterUsed(false);
            int result = chooser.showOpenDialog(null);

            if (result == JFileChooser.APPROVE_OPTION) {
                pptFiles = chooser.getSelectedFiles();
                pref.put("P2E_DEFAULT_PPT_PATH", pptFiles[0].getParentFile().getAbsolutePath());
                List<String> pptFilePaths = Arrays.stream(pptFiles).map(File::getAbsolutePath).collect(Collectors.toList());
                pptFilesText.setText(String.join(",", pptFilePaths));
            }
        });
        excelSelectButton.addActionListener(e -> {

            JFileChooser chooser = new JFileChooser(new File(excelPath));
            chooser.setDialogTitle("Select Output Dir");
            chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            int result = chooser.showOpenDialog(null);

            if (result == JFileChooser.APPROVE_OPTION) {
                excelFileDir = chooser.getSelectedFile();
                pref.put("P2E_DEFAULT_EXCEL_PATH", excelFileDir.getAbsolutePath());
                excelFilesText.setText(excelFileDir.getAbsolutePath());
            }
        });
        generateButton.addActionListener(e -> {
            resultTxt.setText("Generating Excel from the ppt files...\nPlease wait...");
            excelSelectButton.setEnabled(false);
            pptSelectButton.setEnabled(false);
            generateButton.setEnabled(false);
            startThread(resultTxt, pptFiles, excelFileDir, excelSelectButton, pptSelectButton, generateButton);
        });
    }

    private static void startThread(JTextArea resultTxt, File[] pptFiles, File excelFileDir, JButton excelSelectButton, JButton pptSelectButton, JButton generateButton) {

        SwingWorker sw1 = new SwingWorker() {
            // Method
            @Override
            protected String doInBackground()
                    throws Exception {

                Converter converter = new Converter();
                try {
                    String outputFileName = converter.convert(pptFiles, excelFileDir);
                    return outputFileName;

                } catch (IOException ex) {
                    resultTxt.setText("Generate failed: " + ex.getMessage());
                }

                String res = "Finished Execution";
                return res;
            }

            // Method
            @Override
            protected void done() {
                // this method is called when the background
                // thread finishes execution
                try {
                    String outputFileName = (String) get();
                    resultTxt.setText("Generate successfully!\nPlease find the output file at: " + outputFileName);
                    excelSelectButton.setEnabled(true);
                    pptSelectButton.setEnabled(true);
                    generateButton.setEnabled(true);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                } catch (ExecutionException e) {
                    e.printStackTrace();
                }
            }
        };

        // Executes the swingworker on worker thread
        sw1.execute();
    }
}
