window.SP_SYNC = {
  async syncCoverage(mappings, config, dryRun){
    console.log('Would sync coverage to SharePoint', {mappings, config, dryRun});
    alert('Simulation synchronisation Coverage → SharePoint (test local).');
  },
  async syncHazards(hazards, config, dryRun){
    console.log('Would sync hazards to SharePoint', {hazards, config, dryRun});
    alert('Simulation synchronisation Hazards → SharePoint (test local).');
  }
};
