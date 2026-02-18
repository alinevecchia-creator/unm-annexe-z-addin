<script>
// Fonction de commande pour Word Online : ouvre le volet de tâches
// Référencée par onAction="openPane" dans le manifest
Office.onReady(() => {
  // Associe l'action nommée "openPane" à une fonction qui ouvre le task pane
  Office.actions.associate("openPane", async () => {
    try {
      await Office.addin.showAsTaskpane();
    } catch (e) {
      console.error("openPane error:", e);
    }
  });
});
</script>