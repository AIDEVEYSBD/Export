const express = require('express');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();
const port = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());

// Initialize Supabase client
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_ANON_KEY
);

// Utility function to format framework name for display
const formatFrameworkName = (name) => {
  return name.replace(/[_-]/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
};

// Main endpoint for generating framework mapping reports
app.post('/api/framework-mappings/report', async (req, res) => {
  try {
    const { frameworks } = req.body;
    
    if (!frameworks || !Array.isArray(frameworks) || frameworks.length < 2) {
      return res.status(400).json({ 
        error: 'Please provide at least 2 framework names in the frameworks array' 
      });
    }

    console.log('Generating report for frameworks:', frameworks);

    // Get framework IDs and details
    const { data: frameworkData, error: frameworkError } = await supabase
      .from('frameworks')
      .select('id, name')
      .in('name', frameworks);

    if (frameworkError) {
      console.error('Framework query error:', frameworkError);
      return res.status(500).json({ error: 'Failed to fetch frameworks' });
    }

    if (frameworkData.length !== frameworks.length) {
      const foundFrameworks = frameworkData.map(f => f.name);
      const missingFrameworks = frameworks.filter(f => !foundFrameworks.includes(f));
      return res.status(404).json({ 
        error: `Frameworks not found: ${missingFrameworks.join(', ')}` 
      });
    }

    const frameworkIds = frameworkData.map(f => f.id);
    const frameworkLookup = frameworkData.reduce((acc, f) => {
      acc[f.id] = f.name;
      return acc;
    }, {});

    // Get all controls for the specified frameworks
    const { data: controls, error: controlsError } = await supabase
      .from('controls')
      .select('*')
      .in('framework_id', frameworkIds);

    if (controlsError) {
      console.error('Controls query error:', controlsError);
      return res.status(500).json({ error: 'Failed to fetch controls' });
    }

    // Get mappings between controls of these frameworks
    const controlIds = controls.map(c => c.id);
    const { data: mappings, error: mappingsError } = await supabase
      .from('framework_mappings')
      .select(`
        id,
        source_control_id,
        target_control_id,
        mapping_score,
        status,
        explanation,
        created_at
      `)
      .in('source_control_id', controlIds)
      .in('target_control_id', controlIds);

    if (mappingsError) {
      console.error('Mappings query error:', mappingsError);
      return res.status(500).json({ error: 'Failed to fetch mappings' });
    }

    // Create control lookup for easy access
    const controlLookup = controls.reduce((acc, control) => {
      acc[control.id] = control;
      return acc;
    }, {});

    // Generate Excel report
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'GRC Mapping System';
    workbook.lastModifiedBy = 'GRC API';
    workbook.created = new Date();

    // Define professional color scheme
    const colors = {
      primaryBlue: 'FF1E3A8A',
      lightBlue: 'FFE0F2FE', 
      darkGray: 'FF374151',
      lightGray: 'FFF9FAFB',
      green: 'FF059669',
      amber: 'FFD97706',
      red: 'FFDC2626'
    };

    // Sheet 1: Executive Summary
    const summarySheet = workbook.addWorksheet('Executive Summary');
    
    // Set column widths
    summarySheet.columns = [
      { width: 25 }, { width: 20 }, { width: 15 }, { width: 15 }, { width: 20 }
    ];

    // Title
    const titleRow = summarySheet.addRow(['GRC Framework Mapping Report']);
    titleRow.getCell(1).font = { bold: true, size: 18, color: { argb: colors.primaryBlue } };
    titleRow.height = 30;
    summarySheet.mergeCells('A1:E1');

    // Generated timestamp
    const timestampRow = summarySheet.addRow([`Generated: ${new Date().toLocaleString()}`]);
    timestampRow.getCell(1).font = { italic: true, size: 10 };
    summarySheet.mergeCells('A2:E2');

    summarySheet.addRow([]); // Empty row

    // Framework Overview
    const overviewHeaderRow = summarySheet.addRow(['Framework Overview']);
    overviewHeaderRow.getCell(1).font = { bold: true, size: 14, color: { argb: colors.darkGray } };
    summarySheet.mergeCells('A4:E4');

    const overviewHeaders = summarySheet.addRow(['Framework', 'Total Controls', 'Mapped Controls', 'Coverage %', 'Status']);
    overviewHeaders.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.primaryBlue } };
      cell.alignment = { horizontal: 'center' };
    });

    // Calculate framework statistics
    frameworkData.forEach(framework => {
      const frameworkControls = controls.filter(c => c.framework_id === framework.id);
      const mappedControls = mappings.filter(m => 
        frameworkControls.some(fc => fc.id === m.source_control_id) ||
        frameworkControls.some(fc => fc.id === m.target_control_id)
      );
      
      const uniqueMappedControls = new Set();
      mappedControls.forEach(m => {
        if (frameworkControls.some(fc => fc.id === m.source_control_id)) {
          uniqueMappedControls.add(m.source_control_id);
        }
        if (frameworkControls.some(fc => fc.id === m.target_control_id)) {
          uniqueMappedControls.add(m.target_control_id);
        }
      });

      const coverage = frameworkControls.length > 0 ? 
        Math.round((uniqueMappedControls.size / frameworkControls.length) * 100) : 0;
      
      let status = 'Excellent';
      let statusColor = colors.green;
      if (coverage < 70) {
        status = 'Needs Attention';
        statusColor = colors.red;
      } else if (coverage < 90) {
        status = 'Good';
        statusColor = colors.amber;
      }

      const row = summarySheet.addRow([
        formatFrameworkName(framework.name),
        frameworkControls.length,
        uniqueMappedControls.size,
        `${coverage}%`,
        status
      ]);

      row.getCell(5).font = { color: { argb: statusColor }, bold: true };
    });

    // Sheet 2: Detailed Mappings
    const detailSheet = workbook.addWorksheet('Detailed Mappings');
    
    detailSheet.columns = [
      { header: 'Source Framework', key: 'sourceFramework', width: 20 },
      { header: 'Source Control ID', key: 'sourceId', width: 15 },
      { header: 'Source Domain', key: 'sourceDomain', width: 20 },
      { header: 'Source Sub-Domain', key: 'sourceSubDomain', width: 20 },
      { header: 'Source Control', key: 'sourceControl', width: 40 },
      { header: 'Target Framework', key: 'targetFramework', width: 20 },
      { header: 'Target Control ID', key: 'targetId', width: 15 },
      { header: 'Target Domain', key: 'targetDomain', width: 20 },
      { header: 'Target Sub-Domain', key: 'targetSubDomain', width: 20 },
      { header: 'Target Control', key: 'targetControl', width: 40 },
      { header: 'Mapping Score', key: 'mappingScore', width: 12 },
      { header: 'Status', key: 'status', width: 15 },
      { header: 'Explanation', key: 'explanation', width: 50 }
    ];

    // Style the header row
    const headerRow = detailSheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.primaryBlue } };
      cell.alignment = { horizontal: 'center', wrapText: true };
    });
    headerRow.height = 40;

    // Add mapping data
    mappings.forEach((mapping, index) => {
      const sourceControl = controlLookup[mapping.source_control_id];
      const targetControl = controlLookup[mapping.target_control_id];
      
      if (sourceControl && targetControl) {
        const row = detailSheet.addRow({
          sourceFramework: formatFrameworkName(frameworkLookup[sourceControl.framework_id]),
          sourceId: sourceControl.ID,
          sourceDomain: sourceControl.Domain,
          sourceSubDomain: sourceControl['Sub-Domain'],
          sourceControl: sourceControl.Controls,
          targetFramework: formatFrameworkName(frameworkLookup[targetControl.framework_id]),
          targetId: targetControl.ID,
          targetDomain: targetControl.Domain,
          targetSubDomain: targetControl['Sub-Domain'],
          targetControl: targetControl.Controls,
          mappingScore: mapping.mapping_score ? mapping.mapping_score.toFixed(2) : 'N/A',
          status: mapping.status || 'Pending',
          explanation: mapping.explanation || ''
        });

        // Alternate row colors
        if (index % 2 === 0) {
          row.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.lightGray } };
          });
        }

        // Color code mapping scores
        const scoreCell = row.getCell('mappingScore');
        if (mapping.mapping_score) {
          if (mapping.mapping_score >= 0.8) {
            scoreCell.font = { color: { argb: colors.green }, bold: true };
          } else if (mapping.mapping_score >= 0.6) {
            scoreCell.font = { color: { argb: colors.amber }, bold: true };
          } else {
            scoreCell.font = { color: { argb: colors.red }, bold: true };
          }
        }
      }
    });

    // Sheet 3: Unmapped Controls (Gap Analysis)
    const gapSheet = workbook.addWorksheet('Gap Analysis');
    
    gapSheet.columns = [
      { header: 'Framework', key: 'framework', width: 20 },
      { header: 'Control ID', key: 'controlId', width: 15 },
      { header: 'Domain', key: 'domain', width: 20 },
      { header: 'Sub-Domain', key: 'subDomain', width: 20 },
      { header: 'Control Text', key: 'controlText', width: 50 },
      { header: 'Risk Level', key: 'riskLevel', width: 15 },
      { header: 'Recommendation', key: 'recommendation', width: 40 }
    ];

    // Style header
    const gapHeaderRow = gapSheet.getRow(1);
    gapHeaderRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.red } };
      cell.alignment = { horizontal: 'center', wrapText: true };
    });
    gapHeaderRow.height = 40;

    // Find unmapped controls
    const mappedControlIds = new Set();
    mappings.forEach(m => {
      mappedControlIds.add(m.source_control_id);
      mappedControlIds.add(m.target_control_id);
    });

    const unmappedControls = controls.filter(c => !mappedControlIds.has(c.id));

    unmappedControls.forEach((control, index) => {
      const row = gapSheet.addRow({
        framework: formatFrameworkName(frameworkLookup[control.framework_id]),
        controlId: control.ID,
        domain: control.Domain,
        subDomain: control['Sub-Domain'],
        controlText: control.Controls,
        riskLevel: 'High',
        recommendation: 'Requires mapping assessment for compliance coverage'
      });

      // Alternate row colors
      if (index % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.lightGray } };
        });
      }

      // Highlight risk level
      row.getCell('riskLevel').font = { color: { argb: colors.red }, bold: true };
    });

    // Generate buffer and send response
    const buffer = await workbook.xlsx.writeBuffer();
    
    const filename = `Framework_Mapping_Report_${frameworks.join('_vs_')}_${new Date().toISOString().split('T')[0]}.xlsx`;
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);
    
    res.send(buffer);

  } catch (error) {
    console.error('Error generating report:', error);
    res.status(500).json({ error: 'Internal server error generating report' });
  }
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error('Unhandled error:', error);
  res.status(500).json({ error: 'Internal server error' });
});

app.listen(port, () => {
  console.log(`GRC Mapping API running on port ${port}`);
});

module.exports = app;