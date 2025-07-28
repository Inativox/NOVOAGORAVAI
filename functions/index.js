//
// ESTE É O CONTEÚDO CORRETO PARA O ARQUIVO 'functions/index.js'
//

const { onSchedule } = require("firebase-functions/v2/scheduler");
const { logger } = require("firebase-functions");
const admin = require("firebase-admin");

admin.initializeApp();
const db = admin.firestore();

// Esta é a nova sintaxe (V2) para funções agendadas
exports.deletarColecaoAgendada = onSchedule({
  schedule: "every day 20:00",
  timeZone: "America/Sao_Paulo",
  // Aumenta o tempo máximo que a função pode rodar, para garantir que coleções grandes sejam deletadas
  timeoutSeconds: 540, 
  // Aumenta a memória disponível, se necessário para operações intensas
  memory: "1GiB",
}, async (event) => {
  logger.log("Iniciando a exclusão agendada da coleção 'cnpjs_armazenados'.");
  
  const collectionRef = db.collection("cnpjs_armazenados");
  await deleteCollection(collectionRef, 400); // Apaga em lotes de 400

  logger.log("Exclusão da coleção 'cnpjs_armazenados' concluída com sucesso.");
  return null;
});

// --- Funções auxiliares para deletar a coleção em lotes (sem alteração) ---

async function deleteCollection(collectionRef, batchSize) {
  const query = collectionRef.orderBy("__name__").limit(batchSize);

  return new Promise((resolve, reject) => {
    deleteQueryBatch(query, resolve, reject);
  });
}

async function deleteQueryBatch(query, resolve, reject) {
  try {
    const snapshot = await query.get();

    const batchSize = snapshot.size;
    if (batchSize === 0) {
      resolve();
      return;
    }

    const batch = db.batch();
    snapshot.docs.forEach((doc) => {
      batch.delete(doc.ref);
    });
    await batch.commit();

    process.nextTick(() => {
      deleteQueryBatch(query, resolve, reject);
    });

  } catch(err) {
    reject(err);
  }
}