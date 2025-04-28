/*
 * Copyright 2025 [Tu Nombre o Nombre de tu Sitio/Empresa]. Todos los derechos reservados.
 * Script para la Calculadora de Materiales Tablayeso.
 * Maneja la lógica de agregar ítems, calcular materiales y generar reportes.
 * Implementa el criterio de cálculo v2.0, con nombres específicos para Durock Calibre 20 y lógica de tornillos de 1".
 * Implementa lógica de cálculo de paneles con acumuladores fraccionarios/redondeados para áreas pequeñas/grandes (según criterio de imagen).
 * Implementa selección de tipo de panel por cara de muro y para cielos.
 * Implementa Múltiples Entradas de Medida (Segmentos) para Muros Y Cielos.
 * Ajusta el orden de las entradas.
 */

document.addEventListener('DOMContentLoaded', () => {
    const itemsContainer = document.getElementById('items-container');
    const addItemBtn = document.getElementById('add-item-btn');
    const calculateBtn = document.getElementById('calculate-btn');
    const resultsContent = document.getElementById('results-content');
    const downloadOptionsDiv = document.querySelector('.download-options');
    const generatePdfBtn = document.getElementById('generate-pdf-btn');
    const generateExcelBtn = document.getElementById('generate-excel-btn');

    let itemCounter = 0; // To give unique IDs to item blocks

    // Variables to store the last calculated state (needed for PDF/Excel)
    let lastCalculatedTotalMaterials = {}; // Stores final rounded totals for all materials
    let lastCalculatedItemsSpecs = []; // Specs of items included in calculation
    let lastErrorMessages = []; // Store errors as an array of strings

    // --- Constants ---
    const PANEL_RENDIMIENTO_M2 = 2.98; // m2 por panel (rendimiento estándar)
    const SMALL_AREA_THRESHOLD_M2 = 1.5; // Umbral para considerar un área "pequeña" (en m2 por cara/área total del SEGMENTO)

    // Definición de tipos de panel permitidos (deben coincidir con las opciones en el HTML)
    const PANEL_TYPES = [
        "Normal",
        "Resistente a la Humedad",
        "Resistente al Fuego",
        "Alta Resistencia",
        "Exterior" // Asociado comúnmente con Durock, pero aplicable si se usa ese tipo de panel en yeso especial
    ];

     // --- Helper Function for Rounding Up Final Units (Applies per item material quantity, EXCEPT panels in accumulators) ---
    const roundUpFinalUnit = (num) => Math.ceil(num);

    // --- Helper Function to get display name for item type ---
    const getItemTypeName = (typeValue) => {
        switch (typeValue) {
            case 'muro': return 'Muro';
            case 'cielo': return 'Cielo Falso';
            default: return 'Ítem Desconocido';
        }
    };

     // Helper to map item type internal value to a more descriptive name for inputs
     const getItemTypeDescription = (typeValue) => {
         switch (typeValue) {
             case 'muro': return 'Muro';
             case 'cielo': return 'Cielo Falso';
             default: return 'Ítem';
         }
     };


    // --- Helper Function to get the unit for a given material name ---
    const getMaterialUnit = (materialName) => {
         // Map specific names to units based on the new criterion
        // Material names can now include panel types, e.g., "Paneles de Normal"
        if (materialName.startsWith('Paneles de ')) return 'Und'; // Handle all panel types

        switch (materialName) {
            case 'Postes': return 'Und';
            case 'Postes Calibre 20': return 'Und';
            case 'Canales': return 'Und';
            case 'Canales Calibre 20': return 'Und';
            case 'Pasta': return 'Caja';
            case 'Cinta de Papel': return 'm';
            case 'Lija Grano 120': return 'Pliego';
            case 'Clavos con Roldana': return 'Und';
            case 'Fulminantes': return 'Und';
            case 'Tornillos de 1" punta fina': return 'Und';
            case 'Tornillos de 1/2" punta fina': return 'Und';
            case 'Canal Listón': return 'Und';
            case 'Canal Soporte': return 'Und';
            case 'Angular de Lámina': return 'Und';
            case 'Tornillos de 1" punta broca': return 'Und';
            case 'Tornillos de 1/2" punta broca': return 'Und';
            case 'Patas': return 'Und';
            case 'Canal Listón (para cuelgue)': return 'Und';
            case 'Basecoat': return 'Saco'; // Associated with Durock-like panels
            case 'Cinta malla': return 'm'; // Associated with Durock-like panels
            default: return 'Und'; // Default unit if not specified
        }
    };

    // Helper function to get the associated finishing materials based on panel type
    const getFinishingMaterials = (panelType) => {
         const finishing = {};
         // Associate finishing materials based on the panel type name or a category derived from it
         if (panelType === 'Normal' || panelType === 'Resistente a la Humedad' || panelType === 'Resistente al Fuego' || panelType === 'Alta Resistencia') {
             finishing['Pasta'] = 0;
             finishing['Cinta de Papel'] = 0;
             finishing['Lija Grano 120'] = 0;
             finishing['Tornillos de 1" punta fina'] = 0; // Yeso type screws
             finishing['Tornillos de 1/2" punta fina'] = 0; // Yeso type screws for structure
         } else if (panelType === 'Exterior') { // Assuming 'Exterior' implies Durock or similar
             finishing['Basecoat'] = 0;
             finishing['Cinta malla'] = 0;
             finishing['Tornillos de 1" punta broca'] = 0; // Durock type screws
             finishing['Tornillos de 1/2" punta broca'] = 0; // Durock type screws for structure
         }
         return finishing;
    };


    // --- Function to Populate Panel Type Selects ---
    const populatePanelTypes = (selectElement, selectedValue = 'Normal') => {
        selectElement.innerHTML = ''; // Clear existing options
        PANEL_TYPES.forEach(type => {
            const option = document.createElement('option');
            option.value = type;
            option.textContent = type;
            if (type === selectedValue) {
                option.selected = true;
            }
            selectElement.appendChild(option);
        });
    };

     // --- Function to Create a Muro Segment Input Block ---
    const createMuroSegmentBlock = (itemId, segmentNumber) => {
        const segmentHtml = `
            <div class="muro-segment" data-segment-id="${itemId}-mseg-${segmentNumber}">
                 <h4>Segmento ${segmentNumber}</h4>
                 <button type="button" class="remove-segment-btn">X</button>
                <div class="input-group">
                    <label for="mwidth-${itemId}-mseg-${segmentNumber}">Ancho (m):</label>
                    <input type="number" class="item-width" id="mwidth-${itemId}-mseg-${segmentNumber}" step="0.01" min="0" value="3.0">
                </div>
                <div class="input-group">
                    <label for="mheight-${itemId}-mseg-${segmentNumber}">Alto (m):</label>
                    <input type="number" class="item-height" id="mheight-${itemId}-mseg-${segmentNumber}" step="0.01" min="0" value="2.4">
                </div>
            </div>
        `;
        const newElement = document.createElement('div');
        newElement.innerHTML = segmentHtml.trim();
        const segmentBlock = newElement.firstChild;

        // Add remove listener
        const removeButton = segmentBlock.querySelector('.remove-segment-btn');
         removeButton.addEventListener('click', () => {
            const segmentsContainer = segmentBlock.closest('.segments-list'); // Correct selector
            if (segmentsContainer.querySelectorAll('.muro-segment').length > 1) {
                 segmentBlock.remove();
                 // Re-number segments visually after removal
                 segmentsContainer.querySelectorAll('.muro-segment h4').forEach((h4, index) => {
                    h4.textContent = `Segmento ${index + 1}`;
                 });
                 // Clear results and hide download buttons after removal
                 resultsContent.innerHTML = '<p>Segmento eliminado. Recalcula los materiales totales.</p>';
                 downloadOptionsDiv.classList.add('hidden');
                 lastCalculatedTotalMaterials = {};
                 lastCalculatedItemsSpecs = [];
                 lastErrorMessages = [];
            } else {
                 alert("Un muro debe tener al menos un segmento.");
            }
         });

        return segmentBlock;
    };

     // --- Function to Create a Cielo Segment Input Block ---
     const createCieloSegmentBlock = (itemId, segmentNumber) => {
         const segmentHtml = `
            <div class="cielo-segment" data-segment-id="${itemId}-cseg-${segmentNumber}">
                 <h4>Segmento ${segmentNumber}</h4>
                 <button type="button" class="remove-segment-btn">X</button>
                <div class="input-group">
                    <label for="cwidth-${itemId}-cseg-${segmentNumber}">Ancho (m):</label>
                    <input type="number" class="item-width" id="cwidth-${itemId}-cseg-${segmentNumber}" step="0.01" min="0" value="3.0">
                </div>
                <div class="input-group">
                    <label for="clength-${itemId}-cseg-${segmentNumber}">Largo (m):</label>
                    <input type="number" class="item-length" id="clength-${itemId}-cseg-${segmentNumber}" step="0.01" min="0" value="4.0">
                </div>
            </div>
        `;
        const newElement = document.createElement('div');
        newElement.innerHTML = segmentHtml.trim();
        const segmentBlock = newElement.firstChild;

         // Add remove listener
         const removeButton = segmentBlock.querySelector('.remove-segment-btn');
         removeButton.addEventListener('click', () => {
             const segmentsContainer = segmentBlock.closest('.segments-list'); // Correct selector
             if (segmentsContainer.querySelectorAll('.cielo-segment').length > 1) {
                  segmentBlock.remove();
                  // Re-number segments visually after removal
                  segmentsContainer.querySelectorAll('.cielo-segment h4').forEach((h4, index) => {
                     h4.textContent = `Segmento ${index + 1}`;
                  });
                  // Clear results and hide download buttons after removal
                  resultsContent.innerHTML = '<p>Segmento eliminado. Recalcula los materiales totales.</p>';
                  downloadOptionsDiv.classList.add('hidden');
                  lastCalculatedTotalMaterials = {};
                  lastCalculatedItemsSpecs = [];
                  lastErrorMessages = [];
             } else {
                  alert("Un cielo falso debe tener al menos un segmento.");
             }
         });

         return segmentBlock;
     };


     // --- Function to Update Input Visibility WITHIN an Item Block ---
    const updateItemInputVisibility = (itemBlock) => {
        const structureTypeSelect = itemBlock.querySelector('.item-structure-type');
        // Common input groups (some are hidden/shown based on type)
        const facesInputGroup = itemBlock.querySelector('.item-faces-input');
        const muroPanelTypesDiv = itemBlock.querySelector('.muro-panel-types');
        const cieloPanelTypeDiv = itemBlock.querySelector('.cielo-panel-type');
        const postSpacingInputGroup = itemBlock.querySelector('.item-post-spacing-input');
        const doubleStructureInputGroup = itemBlock.querySelector('.item-double-structure-input');
        const plenumInputGroup = itemBlock.querySelector('.item-plenum-input');

        // Type-specific dimension/segment containers
        const muroSegmentsContainer = itemBlock.querySelector('.muro-segments');
        const cieloSegmentsContainer = itemBlock.querySelector('.cielo-segments'); // New container for cielo segments


        const type = structureTypeSelect.value;

        // Reset visibility for ALL type-specific input groups within this block
        facesInputGroup.classList.add('hidden');
        muroPanelTypesDiv.classList.add('hidden');
        cieloPanelTypeDiv.classList.add('hidden');
        postSpacingInputGroup.classList.add('hidden');
        doubleStructureInputGroup.classList.add('hidden');
        plenumInputGroup.classList.add('hidden');
        muroSegmentsContainer.classList.add('hidden'); // Hide muro segments container
        cieloSegmentsContainer.classList.add('hidden'); // Hide cielo segments container


        // Set visibility based on selected type for THIS block
        if (type === 'muro') {
            facesInputGroup.classList.remove('hidden');
            muroPanelTypesDiv.classList.remove('hidden'); // Show wall panel type selectors
            postSpacingInputGroup.classList.remove('hidden'); // Post spacing applies to walls
            doubleStructureInputGroup.classList.remove('hidden'); // Double structure applies to walls
            muroSegmentsContainer.classList.remove('hidden'); // Show muro segments container

             // Hide cielo-specific inputs
            // cieloPanelTypeDiv is already hidden
            // plenumInputGroup is already hidden
            // cieloSegmentsContainer is already hidden

             // Update visibility of face-specific panel type selectors based on faces input
             const facesInput = itemBlock.querySelector('.item-faces');
             const cara2PanelTypeGroup = itemBlock.querySelector('.cara-2-panel-type-group');

             if (parseInt(facesInput.value) === 2) {
                 cara2PanelTypeGroup.classList.remove('hidden');
             } else {
                 cara2PanelTypeGroup.classList.add('hidden');
             }


        } else if (type === 'cielo') {
            cieloPanelTypeDiv.classList.remove('hidden'); // Show ceiling panel type selector
            plenumInputGroup.classList.remove('hidden');
            cieloSegmentsContainer.classList.remove('hidden'); // Show cielo segments container


            // Hide wall-specific inputs for ceiling
            // facesInputGroup is already hidden
            // muroPanelTypesDiv is already hidden
            // postSpacingInputGroup is already hidden
            // doubleStructureInputGroup is already hidden
            // muroSegmentsContainer is already hidden

        }
        // No need for 'else' as all are hidden by default initially
    };

    // --- Function to Create an Item Input Block ---
    const createItemBlock = () => {
        itemCounter++;
        const itemId = `item-${itemCounter}`;

        // Restructured HTML template
        const itemHtml = `
            <div class="item-block" data-item-id="${itemId}">
                <h3>${getItemTypeDescription('muro')} #${itemCounter}</h3>
                <button class="remove-item-btn">Eliminar</button>

                <div class="input-group">
                    <label for="type-${itemId}">Tipo de Estructura:</label>
                    <select class="item-structure-type" id="type-${itemId}">
                        <option value="muro">Muro</option>
                        <option value="cielo">Cielo Falso</option>
                    </select>
                </div>

                <div class="input-group item-faces-input">
                    <label for="faces-${itemId}">Nº de Caras (1 o 2):</label>
                    <input type="number" class="item-faces" id="faces-${itemId}" step="1" min="1" max="2" value="1">
                </div>

                <div class="muro-panel-types">
                    <div class="input-group cara-1-panel-type-group">
                        <label for="cara1-panel-type-${itemId}">Panel Cara 1:</label>
                        <select class="item-cara1-panel-type" id="cara1-panel-type-${itemId}"></select>
                    </div>
                    <div class="input-group cara-2-panel-type-group hidden">
                        <label for="cara2-panel-type-${itemId}">Panel Cara 2:</label>
                        <select class="item-cara2-panel-type" id="cara2-panel-type-${itemId}"></select>
                    </div>
                </div>

                <div class="input-group item-post-spacing-input">
                    <label for="post-spacing-${itemId}">Espaciamiento Postes (m):</label>
                    <input type="number" class="item-post-spacing" id="post-spacing-${itemId}" step="0.01" min="0.1" value="0.40">
                </div>

                <div class="input-group item-double-structure-input">
                    <label for="double-structure-${itemId}">Estructura Doble:</label>
                    <input type="checkbox" class="item-double-structure" id="double-structure-${itemId}">
                </div>


                 <div class="input-group cielo-panel-type hidden">
                    <label for="cielo-panel-type-${itemId}">Tipo de Panel:</label>
                    <select class="item-cielo-panel-type" id="cielo-panel-type-${itemId}"></select>
                </div>

                <div class="input-group item-plenum-input hidden">
                    <label for="plenum-${itemId}">Pleno del Cielo (m):</label>
                    <input type="number" class="item-plenum" id="plenum-${itemId}" step="0.01" min="0" value="0.5">
                </div>

                <div class="muro-segments">
                    <h4>Segmentos del Muro:</h4>
                     <div class="segments-list">
                         </div>
                    <button type="button" class="add-segment-btn">Agregar Segmento</button>
                </div>

                 <div class="cielo-segments hidden">
                    <h4>Segmentos del Cielo Falso:</h4>
                     <div class="segments-list">
                         </div>
                    <button type="button" class="add-segment-btn">Agregar Segmento</button>
                 </div>

            </div>
        `;

        const newElement = document.createElement('div');
        newElement.innerHTML = itemHtml.trim();
        const itemBlock = newElement.firstChild; // Get the actual div element

        itemsContainer.appendChild(itemBlock);


        // Add an initial segment based on the DEFAULT type ('muro')
        const muroSegmentsListContainer = itemBlock.querySelector('.muro-segments .segments-list');
        if (muroSegmentsListContainer) {
             muroSegmentsListContainer.appendChild(createMuroSegmentBlock(itemId, 1));

             // Add listener for "Agregar Segmento" button for muro
             const addMuroSegmentBtn = itemBlock.querySelector('.muro-segments .add-segment-btn');
             addMuroSegmentBtn.addEventListener('click', () => {
                 const currentSegments = muroSegmentsListContainer.querySelectorAll('.muro-segment').length;
                 muroSegmentsListContainer.appendChild(createMuroSegmentBlock(itemId, currentSegments + 1));
                 // Clear results and hide download buttons after adding a segment
                 resultsContent.innerHTML = '<p>Segmento de muro agregado. Recalcula los materiales totales.</p>';
                 downloadOptionsDiv.classList.add('hidden');
                 lastCalculatedTotalMaterials = {};
                 lastCalculatedItemsSpecs = [];
                 lastErrorMessages = [];
             });
        }

         // Add listener for "Agregar Segmento" button for cielo (initially hidden)
         const cieloSegmentsListContainer = itemBlock.querySelector('.cielo-segments .segments-list');
          if (cieloSegmentsListContainer) {
             const addCieloSegmentBtn = itemBlock.querySelector('.cielo-segments .add-segment-btn');
             addCieloSegmentBtn.addEventListener('click', () => {
                 const currentSegments = cieloSegmentsListContainer.querySelectorAll('.cielo-segment').length;
                 cieloSegmentsListContainer.appendChild(createCieloSegmentBlock(itemId, currentSegments + 1));
                 // Clear results and hide download buttons after adding a segment
                 resultsContent.innerHTML = '<p>Segmento de cielo agregado. Recalcula los materiales totales.</p>';
                 downloadOptionsDiv.classList.add('hidden');
                 lastCalculatedTotalMaterials = {};
                 lastCalculatedItemsSpecs = [];
                 lastErrorMessages = [];
             });
          }


        // Populate panel type selects in the new block
        const cara1PanelSelect = itemBlock.querySelector('.item-cara1-panel-type');
        const cara2PanelSelect = itemBlock.querySelector('.item-cara2-panel-type');
        const cieloPanelSelect = itemBlock.querySelector('.item-cielo-panel-type');
        if(cara1PanelSelect) populatePanelTypes(cara1PanelSelect);
        if(cara2PanelSelect) populatePanelTypes(cara2PanelSelect);
        if(cieloPanelSelect) populatePanelTypes(cieloPanelSelect);


        // Add event listener to the new select element IN THIS BLOCK
        const structureTypeSelect = itemBlock.querySelector('.item-structure-type');
        structureTypeSelect.addEventListener('change', (event) => {
            const selectedType = event.target.value;
            // Update the h3 title based on the selected type
            itemBlock.querySelector('h3').textContent = `${getItemTypeDescription(selectedType)} #${itemCounter}`;

            // Clear existing segments when changing type
            const muroSegmentsList = itemBlock.querySelector('.muro-segments .segments-list');
            const cieloSegmentsList = itemBlock.querySelector('.cielo-segments .segments-list');

            if (selectedType === 'muro') {
                // Clear cielo segments and add a muro segment if needed
                if (cieloSegmentsList) cieloSegmentsList.innerHTML = '';
                if (muroSegmentsList && muroSegmentsList.querySelectorAll('.muro-segment').length === 0) {
                     muroSegmentsList.appendChild(createMuroSegmentBlock(itemId, 1));
                }
            } else if (selectedType === 'cielo') {
                 // Clear muro segments and add a cielo segment if needed
                 if (muroSegmentsList) muroSegmentsList.innerHTML = '';
                 if (cieloSegmentsList && cieloSegmentsList.querySelectorAll('.cielo-segment').length === 0) {
                     cieloSegmentsList.appendChild(createCieloSegmentBlock(itemId, 1));
                 }
            }


            updateItemInputVisibility(itemBlock);
            // Clear results and hide download buttons on type change
             resultsContent.innerHTML = '<p>Tipo de ítem cambiado. Recalcula los materiales totales.</p>';
             downloadOptionsDiv.classList.add('hidden');
             lastCalculatedTotalMaterials = {};
             lastCalculatedItemsSpecs = [];
             lastErrorMessages = [];
        });

         // Add event listener to faces input to toggle visibility of Cara 2 panel type
        const facesInput = itemBlock.querySelector('.item-faces');
        if(facesInput) { // Only add listener if faces input exists (for muros)
            facesInput.addEventListener('input', () => updateItemInputVisibility(itemBlock));
        }


        // Add event listener to the new remove button
        const removeButton = itemBlock.querySelector('.remove-item-btn');
        removeButton.addEventListener('click', () => {
            itemBlock.remove(); // Remove the block from the DOM
            // Clear results and hide download buttons after removal for immediate feedback
             resultsContent.innerHTML = '<p>Ítem eliminado. Recalcula los materiales totales.</p>';
             downloadOptionsDiv.classList.add('hidden'); // Hide download options
             // Also reset stored data on item removal
             lastCalculatedTotalMaterials = {};
             lastCalculatedItemsSpecs = [];
             lastErrorMessages = [];
             // Re-evaluate if calculate button should be disabled (if no items left)
             toggleCalculateButtonState();
        });


        // Set initial visibility for the inputs in the new block (defaults to muro)
        updateItemInputVisibility(itemBlock);
        // Re-evaluate if calculate button should be enabled (since an item was added)
        toggleCalculateButtonState();

        return itemBlock; // Return the created element
    };

     // --- Function to Enable/Disable Calculate Button ---
    const toggleCalculateButtonState = () => {
        const itemBlocks = itemsContainer.querySelectorAll('.item-block');
        calculateBtn.disabled = itemBlocks.length === 0;
    };


    // --- Main Calculation Function for ALL Items ---
    const calculateMaterials = () => {
        console.log("Iniciando cálculo de materiales...");
        const itemBlocks = itemsContainer.querySelectorAll('.item-block');

        // --- Accumulators for Panels (per panel type) based on Image Logic ---
        let panelAccumulators = {};
         PANEL_TYPES.forEach(type => {
            panelAccumulators[type] = {
                suma_fraccionaria_pequenas: 0.0,
                suma_redondeada_otros: 0
            };
        });
         console.log("Acumuladores de paneles inicializados:", panelAccumulators);

        // --- Accumulator for ALL other materials (rounded per item and summed) ---
        let otherMaterialsTotal = {};

        let currentCalculatedItemsSpecs = []; // Array to store specs of validly calculated items
        let currentErrorMessages = []; // Use an array to collect validation error messages

        // Clear previous results and hide download buttons initially
        resultsContent.innerHTML = '';
        downloadOptionsDiv.classList.add('hidden');

        if (itemBlocks.length === 0) {
            console.log("No hay ítems para calcular.");
            resultsContent.innerHTML = '<p style="color: orange; text-align: center; font-style: italic;">Por favor, agrega al menos un Muro o Cielo para calcular.</p>';
             // Store empty results
             lastCalculatedTotalMaterials = {};
             lastCalculatedItemsSpecs = [];
             lastErrorMessages = ['No hay ítems agregados para calcular.'];
            return;
        }

        console.log(`Procesando ${itemBlocks.length} ítems.`);
        // Iterate through each item block and calculate its materials
        itemBlocks.forEach(itemBlock => {
            const itemNumber = itemBlock.querySelector('h3').textContent.split('#')[1]; // Extract number like "1"
            const type = itemBlock.querySelector('.item-structure-type').value;
            const itemId = itemBlock.dataset.itemId; // Get the unique item ID

            // Get common values (some are type-specific and might be NaN/null/false)
            const facesInput = itemBlock.querySelector('.item-faces');
            const faces = facesInput && !facesInput.closest('.hidden') ? parseInt(facesInput.value) : NaN;

            const plenumInput = itemBlock.querySelector('.item-plenum');
            const plenum = plenumInput && !plenumInput.closest('.hidden') ? parseFloat(plenumInput.value) : NaN;

            const isDoubleStructureInput = itemBlock.querySelector('.item-double-structure');
            const isDoubleStructure = isDoubleStructureInput && !isDoubleStructureInput.closest('.hidden') ? isDoubleStructureInput.checked : false;

            const postSpacingInput = itemBlock.querySelector('.item-post-spacing');
            const postSpacing = postSpacingInput && !postSpacingInput.closest('.hidden') ? parseFloat(postSpacingInput.value) : NaN;

             // Get panel types based on visibility and type
            const cara1PanelTypeSelect = itemBlock.querySelector('.item-cara1-panel-type');
            const cara1PanelType = cara1PanelTypeSelect && !cara1PanelTypeSelect.closest('.hidden') ? cara1PanelTypeSelect.value : null;

            const cara2PanelTypeSelect = itemBlock.querySelector('.item-cara2-panel-type');
            // Only read if faces is 2, selector is visible, and the value is not null/empty
            const cara2PanelType = (faces === 2 && cara2PanelTypeSelect && !cara2PanelTypeSelect.closest('.hidden') && cara2PanelTypeSelect.value) ? cara2PanelTypeSelect.value : null;


            const cieloPanelTypeSelect = itemBlock.querySelector('.item-cielo-panel-type');
            const cieloPanelType = cieloPanelTypeSelect && !cieloPanelTypeSelect.closest('.hidden') ? cieloPanelTypeSelect.value : null;


            console.log(`Procesando Ítem #${itemNumber} (ID: ${itemId}): Tipo=${type}`);

             // Basic Validation for Each Item
             let itemSpecificErrors = [];
             let itemValidatedSpecs = { // Store specs for valid items *before* calculation
                 id: itemId,
                 number: itemNumber,
                 type: type,
                 faces: type === 'muro' ? faces : NaN,   // Only store faces for muros
                 cara1PanelType: type === 'muro' ? cara1PanelType : null, // Only store for muros
                 cara2PanelType: type === 'muro' && faces === 2 ? cara2PanelType : null, // Only store for muros (if 2 faces)
                 cieloPanelType: type === 'cielo' ? cieloPanelType : null, // Only store for cielos
                 postSpacing: type === 'muro' ? postSpacing : NaN, // Only store for muros
                 plenum: type === 'cielo' ? plenum : NaN, // Only store for cielos
                 isDoubleStructure: type === 'muro' ? isDoubleStructure : false, // Only store for muros
                 segments: [] // Array to store valid segments (muro or cielo)
             };


            // Object to hold calculated *other* materials for THIS single item (initial floats)
            // Initialize finishing based on the primary panel type for the item (Cara 1 for Muro, Cielo type for Cielo)
            let itemOtherMaterialsFloat = getFinishingMaterials(type === 'muro' ? cara1PanelType : cieloPanelType);


            // --- Calculation Logic for the CURRENT Item ---

            if (type === 'muro') {
                const segmentBlocks = itemBlock.querySelectorAll('.muro-segment');
                let totalMuroAreaForPanelsFinishing = 0; // Renamed for clarity
                let totalMuroWidthForStructure = 0;
                let hasValidSegment = false; // Flag to check if at least one segment is valid

                 if (segmentBlocks.length === 0) {
                     itemSpecificErrors.push('Muro debe tener al menos un segmento de medida.');
                 } else {
                     segmentBlocks.forEach((segBlock, index) => {
                         const segmentWidth = parseFloat(segBlock.querySelector('.item-width').value); // Use .item-width class
                         const segmentHeight = parseFloat(segBlock.querySelector('.item-height').value); // Use .item-height class
                         const segmentNumber = index + 1;

                         // Validate segment dimensions
                         if (isNaN(segmentWidth) || segmentWidth <= 0 || isNaN(segmentHeight) || segmentHeight <= 0) {
                             itemSpecificErrors.push(`Segmento ${segmentNumber}: Dimensiones inválidas (Ancho y Alto deben ser > 0)`);
                             return; // Skip this segment but continue validating others
                         }

                         // If segment dimensions are valid
                         hasValidSegment = true;
                         const segmentArea = segmentWidth * segmentHeight;
                         totalMuroAreaForPanelsFinishing += segmentArea; // Sum area for panels/finishing
                         totalMuroWidthForStructure += segmentWidth; // Sum width for structure

                         itemValidatedSpecs.segments.push({ // Store valid segment specs
                             number: segmentNumber,
                             width: segmentWidth,
                             height: segmentHeight,
                             area: segmentArea // Store calculated area for report
                         });


                         // --- Panel Calculation for THIS Muro Segment (using Image Logic) ---
                         // Panel calculation is based on the area of *each segment*
                         const panelTypeFace1 = cara1PanelType;
                         const panelesFloatFace1 = segmentArea / PANEL_RENDIMIENTO_M2;

                         if (segmentArea < SMALL_AREA_THRESHOLD_M2 && segmentArea > 0) {
                             panelAccumulators[panelTypeFace1].suma_fraccionaria_pequenas += panelesFloatFace1;
                             console.log(`Muro #${itemNumber} Cara 1 Segmento ${segmentNumber} (${panelTypeFace1}): Área pequeña (${segmentArea.toFixed(2)}m2). Sumando fraccional (${panelesFloatFace1.toFixed(2)}) a acumulador.`);
                         } else if (segmentArea >= SMALL_AREA_THRESHOLD_M2) {
                             const panelesRoundedFace1 = roundUpFinalUnit(panelesFloatFace1);
                             panelAccumulators[panelTypeFace1].suma_redondeada_otros += panelesRoundedFace1;
                              console.log(`Muro #${itemNumber} Cara 1 Segmento ${segmentNumber} (${panelTypeFace1}): Área grande (${segmentArea.toFixed(2)}m2). Sumando redondeado (${panelesRoundedFace1}) a acumulador.`);
                         }


                         if (faces === 2 && cara2PanelType) { // Only calculate for Face 2 if 2 faces are selected and type exists
                             const panelTypeFace2 = cara2PanelType;
                             const panelesFloatFace2 = segmentArea / PANEL_RENDIMIENTO_M2;

                             if (segmentArea < SMALL_AREA_THRESHOLD_M2 && segmentArea > 0) {
                                 panelAccumulators[panelTypeFace2].suma_fraccionaria_pequenas += panelesFloatFace2;
                                 console.log(`Muro #${itemNumber} Cara 2 Segmento ${segmentNumber} (${panelTypeFace2}): Área pequeña (${segmentArea.toFixed(2)}m2). Sumando fraccional (${panelesFloatFace2.toFixed(2)}) a acumulador.`);
                             } else if (segmentArea >= SMALL_AREA_THRESHOLD_M2) {
                                 const panelesRoundedFace2 = roundUpFinalUnit(panelesFloatFace2);
                                 panelAccumulators[panelTypeFace2].suma_redondeada_otros += panelesRoundedFace2;
                                 console.log(`Muro #${itemNumber} Cara 2 Segmento ${segmentNumber} (${panelTypeFace2}): Área grande (${segmentArea.toFixed(2)}m2). Sumando redondeado (${panelesRoundedFace2}) a acumulador.`);
                             }
                         }
                     }); // End of segmentBlocks.forEach

                     // Add validation specific to Muros AFTER processing segments
                     if (!hasValidSegment && segmentBlocks.length > 0) {
                         itemSpecificErrors.push('Ningún segmento de muro tiene dimensiones válidas (> 0).');
                     }
                     if (isNaN(faces) || (faces !== 1 && faces !== 2)) itemSpecificErrors.push('Nº Caras inválido (debe ser 1 o 2)');
                     if (isNaN(postSpacing) || postSpacing <= 0) itemSpecificErrors.push('Espaciamiento Postes inválido (debe ser > 0)');
                     if (!cara1PanelType || !PANEL_TYPES.includes(cara1PanelType)) itemSpecificErrors.push('Tipo de Panel Cara 1 inválido.');
                     // Check cara2PanelType only if faces is 2
                     if (faces === 2 && (!cara2PanelType || !PANEL_TYPES.includes(cara2PanelType))) itemSpecificErrors.push('Tipo de Panel Cara 2 inválido para 2 caras.');
                     // Check if total width and area are > 0 if there was at least one valid segment
                      if (hasValidSegment && (totalMuroWidthForStructure <= 0 || totalMuroAreaForPanelsFinishing <= 0)) {
                           // This error might be redundant if segment validation catches 0, but keep for safety
                           // itemSpecificErrors.push('La suma de anchos y áreas de los segmentos debe ser mayor a 0.');
                      }


                     // Store total calculated values for muros
                     itemValidatedSpecs.totalMuroArea = totalMuroAreaForPanelsFinishing;
                     itemValidatedSpecs.totalMuroWidth = totalMuroWidthForStructure;

                 } // End if segmentBlocks.length > 0


                // --- Other Materials Calculation for THIS Muro Item (Structure, Finishing, Screws) ---
                // Only proceed with material calculation if there are no validation errors *so far* for this item and at least one valid segment
                if (itemSpecificErrors.length === 0 && hasValidSegment) {
                    // Structure calculation is based on the *total accumulated width* of all segments

                    // Postes (based on total width)
                    let postesFloat;
                    if (totalMuroWidthForStructure > 0 && postSpacing > 0) {
                         // Original logic: if total width < post spacing, need 2 postes (start/end). Otherwise, floor(total width / spacing) + 1.
                         if (totalMuroWidthForStructure < postSpacing) { // e.g., width 0.5m, spacing 0.6m
                            postesFloat = 2;
                         } else { // e.g., width 3m, spacing 0.4m -> 3/0.4 = 7.5 -> floor(7.5) + 1 = 8.5 -> need 8 postes minimum, plus ends?
                            // Let's re-evaluate: number of bays is floor(width/spacing). Number of studs is bays + 1.
                            postesFloat = Math.floor(totalMuroWidthForStructure / postSpacing) + 1;
                         }
                    } else {
                        postesFloat = 0;
                    }
                    if (isDoubleStructure) postesFloat *= 2;
                     // Determine Postes type based on panel type of Cara 1
                     itemOtherMaterialsFloat[cara1PanelType === 'Exterior' ? 'Postes Calibre 20' : 'Postes'] = postesFloat;


                    // Canales (based on total width)
                    let canalesFloat;
                     if (totalMuroWidthForStructure > 0) {
                       // Need 2 channels (top/bottom) along the total width. Standard length 3.05m.
                        canalesFloat = (totalMuroWidthForStructure * 2) / 3.05;
                    } else {
                        canalesFloat = 0;
                    }
                    if (isDoubleStructure) canalesFloat *= 2;
                    // Determine Canales type based on panel type of Cara 1
                    itemOtherMaterialsFloat[cara1PanelType === 'Exterior' ? 'Canales Calibre 20' : 'Canales'] = canalesFloat;


                     // Acabado y Tornillos de Panel (based on total accumulated area of ALL segments)
                    const primaryPanelTypeForFinishing = cara1PanelType; // Use Cara 1 type for finishing logic

                    if (primaryPanelTypeForFinishing === 'Normal' || primaryPanelTypeForFinishing === 'Resistente a la Humedad' || primaryPanelTypeForFinishing === 'Resistente al Fuego' || primaryPanelTypeForFinishing === 'Alta Resistencia') {
                        // Pasta (Yeso) - Calculation based on Total Accumulated Area
                        itemOtherMaterialsFloat['Pasta'] = totalMuroAreaForPanelsFinishing > 0 ? totalMuroAreaForPanelsFinishing / 22 : 0; // Area / Rendimiento (22 m2/caja)

                        // Cinta de Papel - Calculation based on Total Accumulated Area
                        itemOtherMaterialsFloat['Cinta de Papel'] = totalMuroAreaForPanelsFinishing > 0 ? totalMuroAreaForPanelsFinishing * 1 : 0; // 1 meter per m2

                        // Lija Grano 120 - Calculation based on Total Accumulated Area
                        // Estimate total panels for the item (total area / 2.98), then divide by 2.
                        // Ensure totalMuroAreaForPanelsFinishing > 0 before dividing
                        itemOtherMaterialsFloat['Lija Grano 120'] = totalMuroAreaForPanelsFinishing > 0 ? (totalMuroAreaForPanelsFinishing / PANEL_RENDIMIENTO_M2) / 2 : 0;


                        // Tornillos de 1" punta fina (for Yeso type panels on this item)
                        // Base this on the estimated *total* panel count for the item's total area.
                        // Ensure totalMuroAreaForPanelsFinishing > 0 before dividing
                        itemOtherMaterialsFloat['Tornillos de 1" punta fina'] = totalMuroAreaForPanelsFinishing > 0 ? (totalMuroAreaForPanelsFinishing / PANEL_RENDIMIENTO_M2) * 40 : 0;


                    } else if (primaryPanelTypeForFinishing === 'Exterior') { // Assuming 'Exterior' implies Durock or similar
                         // Basecoat (Durock) - Calculation based on Total Accumulated Area
                        itemOtherMaterialsFloat['Basecoat'] = totalMuroAreaForPanelsFinishing > 0 ? totalMuroAreaForPanelsFinishing / 8 : 0; // Area / Rendimiento (8 m2/saco)

                        // Cinta malla (Durock) - Calculation based on Total Accumulated Area
                         itemOtherMaterialsFloat['Cinta malla'] = totalMuroAreaForPanelsFinishing > 0 ? totalMuroAreaForPanelsFinishing * 1 : 0; // 1 meter per m2

                        // Tornillos de 1" punta broca (for Durock type panels on this item)
                         // Base on estimated *total* panel count for item's total area.
                         itemOtherMaterialsFloat['Tornillos de 1" punta broca'] = totalMuroAreaForPanelsFinishing > 0 ? (totalMuroAreaForPanelsFinishing / PANEL_RENDIMIENTO_M2) * 40 : 0;
                    }


                     // --- Calculate Tornillos/Clavos Based on Rounded Component Counts for THIS Item ---
                     // These are calculated per item based on rounded structure counts for THIS item.
                     // Only calculate if totalMuroWidthForStructure > 0
                    if (totalMuroWidthForStructure > 0) {
                        let roundedPostes = roundUpFinalUnit(itemOtherMaterialsFloat[cara1PanelType === 'Exterior' ? 'Postes Calibre 20' : 'Postes'] || 0);
                        let roundedCanales = roundUpFinalUnit(itemOtherMaterialsFloat[cara1PanelType === 'Exterior' ? 'Canales Calibre 20' : 'Canales'] || 0);

                        // Clavos con Roldana (8 per Canal) - Use rounded Canales for this item
                         itemOtherMaterialsFloat['Clavos con Roldana'] = roundedCanales * 8;
                         itemOtherMaterialsFloat['Fulminantes'] = itemOtherMaterialsFloat['Clavos con Roldana']; // Igual cantidad

                         // Tornillos 1/2" (4 per Poste) - Use rounded Postes for this item
                        itemOtherMaterialsFloat[cara1PanelType === 'Exterior' ? 'Tornillos de 1/2" punta broca' : 'Tornillos de 1/2" punta fina'] = roundedPostes * 4;
                    }


                } // End if itemSpecificErrors.length === 0 && hasValidSegment


            } else if (type === 'cielo') {
                 const segmentBlocks = itemBlock.querySelectorAll('.cielo-segment');
                 let totalCieloAreaForPanelsFinishing = 0; // Total area from all segments
                 let totalCieloPerimeterForAngular = 0; // Sum of (width + length) for angular calculation
                 let hasValidSegment = false;
                 let validSegments = []; // Temp array to store valid segment specs

                 if (segmentBlocks.length === 0) {
                     itemSpecificErrors.push('Cielo Falso debe tener al menos un segmento de medida.');
                 } else {
                     segmentBlocks.forEach((segBlock, index) => {
                          const segmentWidth = parseFloat(segBlock.querySelector('.item-width').value); // Use .item-width class
                          const segmentLength = parseFloat(segBlock.querySelector('.item-length').value); // Use .item-length class
                          const segmentNumber = index + 1;

                          // Validate segment dimensions
                          if (isNaN(segmentWidth) || segmentWidth <= 0 || isNaN(segmentLength) || segmentLength <= 0) {
                              itemSpecificErrors.push(`Segmento ${segmentNumber}: Dimensiones inválidas (Ancho y Largo deben ser > 0)`);
                              return; // Skip this segment
                          }

                          // If segment dimensions are valid
                          hasValidSegment = true;
                          const segmentArea = segmentWidth * segmentLength;
                          totalCieloAreaForPanelsFinishing += segmentArea; // Sum area
                          totalCieloPerimeterForAngular += (segmentWidth + segmentLength); // Sum (W+L) for angular calculation

                          validSegments.push({ // Store valid segment specs
                             number: segmentNumber,
                             width: segmentWidth,
                             length: segmentLength,
                             area: segmentArea // Store calculated area for report
                         });


                          // --- Panel Calculation for THIS Cielo Segment (using Image Logic) ---
                          // Panel calculation is based on the area of *each segment*
                          const panelTypeCielo = cieloPanelType;
                          const panelesFloatCielo = segmentArea / PANEL_RENDIMIENTO_M2;

                          if (segmentArea < SMALL_AREA_THRESHOLD_M2 && segmentArea > 0) {
                              panelAccumulators[panelTypeCielo].suma_fraccionaria_pequenas += panelesFloatCielo;
                              console.log(`Cielo #${itemNumber} Segmento ${segmentNumber} (${panelTypeCielo}): Área pequeña (${segmentArea.toFixed(2)}m2). Sumando fraccional (${panelesFloatCielo.toFixed(2)}) a acumulador.`);
                          } else if (segmentArea >= SMALL_AREA_THRESHOLD_M2) {
                              const panelesRoundedCielo = roundUpFinalUnit(panelesFloatCielo);
                              panelAccumulators[panelTypeCielo].suma_redondeada_otros += panelesRoundedCielo;
                              console.log(`Cielo #${itemNumber} Segmento ${segmentNumber} (${panelTypeCielo}): Área grande (${segmentArea.toFixed(2)}m2). Sumando redondeado (${panelesRoundedCielo}) a acumulador.`);
                          }
                     }); // End of segmentBlocks.forEach for cielo

                     // Add validation specific to Cielos AFTER processing segments
                     if (!hasValidSegment && segmentBlocks.length > 0) {
                         itemSpecificErrors.push('Ningún segmento de cielo falso tiene dimensiones válidas (> 0).');
                     }
                      // Plenum validation only if visible and required (check plenum input existence and value)
                     const plenumInput = itemBlock.querySelector('.item-plenum'); // Re-get input as it might be hidden
                      if (itemBlock.querySelector('.item-plenum-input') && !itemBlock.querySelector('.item-plenum-input').classList.contains('hidden') && (isNaN(plenum) || plenum < 0)) {
                          itemSpecificErrors.push('Pleno inválido (debe ser >= 0)');
                      }
                     if (!cieloPanelType || !PANEL_TYPES.includes(cieloPanelType)) itemSpecificErrors.push('Tipo de Panel de Cielo inválido.');
                      // Check if total area is > 0 if there was at least one valid segment
                      if (hasValidSegment && totalCieloAreaForPanelsFinishing <= 0) {
                           // This error might be redundant if segment validation catches 0, but keep for safety
                           // itemSpecificErrors.push('La suma de áreas de los segmentos de cielo falso debe ser mayor a 0.');
                      }

                     // Store validated segments and totals for cielos
                     itemValidatedSpecs.segments = validSegments;
                     itemValidatedSpecs.totalCieloArea = totalCieloAreaForPanelsFinishing;
                     itemValidatedSpecs.totalCieloPerimeterSum = totalCieloPerimeterForAngular; // Store sum W+L for angular calc


                 } // End if segmentBlocks.length > 0 for cielo


                 // --- Other Materials Calculation for THIS Cielo Item (Structure, Finishing, Screws) ---
                 // Only proceed if there are no validation errors *so far* for this item and at least one valid segment
                if (itemSpecificErrors.length === 0 && hasValidSegment) {

                     // Canal Listón (based on total area)
                     let canalListonFloat = 0;
                     if (totalCieloAreaForPanelsFinishing > 0) {
                         // Formula seems to be: total area / spacing (0.40) / length of piece (3.66)
                         canalListonFloat = (totalCieloAreaForPanelsFinishing / 0.40) / 3.66;
                     }
                     itemOtherMaterialsFloat['Canal Listón'] = canalListonFloat;


                     // Canal Soporte (based on total area)
                     let canalSoporteFloat = 0;
                     if (totalCieloAreaForPanelsFinishing > 0) {
                         // Formula seems to be: total area / spacing (0.90) / length of piece (3.66)
                        canalSoporteFloat = (totalCieloAreaForPanelsFinishing / 0.90) / 3.66;
                     }
                     itemOtherMaterialsFloat['Canal Soporte'] = canalSoporteFloat;


                     // Angular de Lámina (based on sum of W+L of segments)
                     let angularLaminaFloat = 0;
                     if (totalCieloPerimeterForAngular > 0) {
                          // Sum of W+L of all segments, divided by standard angular length (2.44)
                          // Note: This is a simplification for perimeter estimation with segments
                         angularLaminaFloat = totalCieloPerimeterForAngular / 2.44;
                     }
                     itemOtherMaterialsFloat['Angular de Lámina'] = angularLaminaFloat;


                    // Patas (Soportes) - Calculation based on Canal Soporte quantity for THIS item
                     // Calculate base Patas quantity first, then use its rounded value later for Cuelgue/Screws
                    let patasFloat = (itemOtherMaterialsFloat['Canal Soporte'] || 0) * 4;
                    itemOtherMaterialsFloat['Patas'] = patasFloat; // Store float value


                     // Canal Listón (para cuelgue) - Represents the number of 3.66m profiles needed to cut the hangers
                     // Calculation based on rounded Patas quantity and Plenum for THIS item
                     let roundedPatasForCuelgue = roundUpFinalUnit(patasFloat || 0); // Round patas float value
                     if (plenum > 0 && roundedPatasForCuelgue > 0 && !isNaN(plenum)) {
                        // Number of cuelgues needed = roundedPatasForCuelgue. Each cuelgue has length = plenum. Total length = roundedPatasForCuelgue * plenum. Convert to 3.66m pieces.
                        itemOtherMaterialsFloat['Canal Listón (para cuelgue)'] = (roundedPatasForCuelgue * plenum) / 3.66;
                     } else {
                         itemOtherMaterialsFloat['Canal Listón (para cuelgue)'] = 0;
                     }


                     // Tornillos 1" punta broca (for Cielo panels - assuming Yeso type uses broca for structure attachment to ceiling)
                     // Base this on the estimated *total* panel count for the item's total area.
                     itemOtherMaterialsFloat['Tornillos de 1" punta broca'] = totalCieloAreaForPanelsFinishing > 0 ? (totalCieloAreaForPanelsFinishing / PANEL_RENDIMIENTO_M2) * 40 : 0; // Use broca for ceiling panel attachment


                     // Acabado (Pasta, Cinta de Papel, Lija / Basecoat, Cinta malla) - Same calculation as Muros, based on total accumulated panel area
                     const primaryPanelTypeForFinishing = cieloPanelType; // Use Cielo type for finishing logic

                     if (primaryPanelTypeForFinishing === 'Normal' || primaryPanelTypeForFinishing === 'Resistente a la Humedad' || primaryPanelTypeForFinishing === 'Resistente al Fuego' || primaryPanelTypeForFinishing === 'Alta Resistencia') {
                         // Pasta (Yeso) - Calculation based on Area (total item area)
                         itemOtherMaterialsFloat['Pasta'] = totalCieloAreaForPanelsFinishing > 0 ? totalCieloAreaForPanelsFinishing / 22 : 0;

                         // Cinta de Papel - Calculation based on Area
                         itemOtherMaterialsFloat['Cinta de Papel'] = totalCieloAreaForPanelsFinishing > 0 ? totalCieloAreaForPanelsFinishing * 1 : 0;

                         // Lija Grano 120 - Calculation based on Area (estimated panels / 2)
                         itemOtherMaterialsFloat['Lija Grano 120'] = totalCieloAreaForPanelsFinishing > 0 ? (totalCieloAreaForPanelsFinishing / PANEL_RENDIMIENTO_M2) / 2 : 0;

                     } else if (primaryPanelTypeForFinishing === 'Exterior') { // Assuming 'Exterior' implies Durock or similar
                          // Basecoat (Durock) - Calculation based on Area
                         itemOtherMaterialsFloat['Basecoat'] = totalCieloAreaForPanelsFinishing > 0 ? totalCieloAreaForPanelsFinishing / 8 : 0;

                         // Cinta malla (Durock) - Calculation based on Area
                         itemOtherMaterialsFloat['Cinta malla'] = totalCieloAreaForPanelsFinishing > 0 ? totalCieloAreaForPanelsFinishing * 1 : 0;
                     }

                     // --- Calculate Tornillos/Clavos Based on Rounded Component Counts for THIS Item ---
                     // These are calculated per item based on rounded structure counts for THIS item.
                     // Only calculate if there is area/structure components
                     if (totalCieloAreaForPanelsFinishing > 0 || totalCieloPerimeterForAngular > 0) {
                         let roundedAngularLamina = roundUpFinalUnit(itemOtherMaterialsFloat['Angular de Lámina'] || 0);
                         let roundedCanalSoporte = roundUpFinalUnit(itemOtherMaterialsFloat['Canal Soporte'] || 0);
                         let roundedCanalListon = roundUpFinalUnit(itemOtherMaterialsFloat['Canal Listón'] || 0);
                          // Patas calculation depends on Canal Soporte, round its calculated float value *for this item*
                         let roundedPatas = roundUpFinalUnit(itemOtherMaterialsFloat['Patas'] || 0); // Use the previously stored float value

                         // Clavos con Roldana (5 per Angular + 8 per Canal Soporte) - Use rounded counts for this item
                         itemOtherMaterialsFloat['Clavos con Roldana'] = (roundedAngularLamina * 5) + (roundedCanalSoporte * 8);
                         itemOtherMaterialsFloat['Fulminantes'] = itemOtherMaterialsFloat['Clavos con Roldana']; // Igual cantidad

                         // Tornillos 1/2" punta fina (12 per Canal Listón + 2 per Pata) - Use rounded counts for this item
                         itemOtherMaterialsFloat['Tornillos de 1/2" punta fina'] = (roundedCanalListon * 12) + (roundedPatas * 2); // Cielo structure uses punta fina

                     }


                 } // End if itemSpecificErrors.length === 0 && hasValidSegment for cielo


            } else {
                // Unknown type (shouldn't happen with validation)
                 itemSpecificErrors.push('Tipo de estructura desconocido.');
            }

            console.log(`Ítem #${itemNumber}: Errores de validación - ${itemSpecificErrors.length}`);

            // If item has errors, add to global error list and skip calculation for this item
            if (itemSpecificErrors.length > 0) {
                 const itemTitle = itemBlock.querySelector('h3').textContent;
                 currentErrorMessages.push(`Error en ${itemTitle}: ${itemSpecificErrors.join(', ')}`);
                 console.warn(`Item inválido o incompleto: ${itemTitle}. Errores: ${itemSpecificErrors.join(', ')}. Este ítem no se incluirá en el cálculo total.`);
                 // Do NOT add to currentCalculatedItemsSpecs if there are errors
                 return; // Skip calculation and summing for this invalid item
            }

             // Store the validated specs for this item (segments already pushed for muro/cielo)
             // Only push if there were no errors for this item
             currentCalculatedItemsSpecs.push(itemValidatedSpecs);


            console.log(`Ítem #${itemNumber}: Otros materiales calculados (float) antes de redondear individualmente:`, itemOtherMaterialsFloat);

             // --- Round Up *Other* Material Quantities for THIS Item and Sum to Total ---
            // Panels are handled separately in accumulators.
            // This block executes only for valid items (no errors)
            for (const material in itemOtherMaterialsFloat) {
                if (itemOtherMaterialsFloat.hasOwnProperty(material)) {
                    const floatQuantity = itemOtherMaterialsFloat[material];
                    // Ensure the value is a valid number before rounding and summing
                    // Only sum positive quantities or quantities calculated based on rounded components (which could be 0 if components are 0)
                    if (!isNaN(floatQuantity)) { // Sum all valid numbers, even 0 if it results from calc
                         const roundedQuantity = roundUpFinalUnit(floatQuantity);
                         // Sum this rounded quantity from the current item to the overall total of OTHER materials
                         otherMaterialsTotal[material] = (otherMaterialsTotal[material] || 0) + roundedQuantity;
                    }
                }
            }
             console.log(`Ítem #${itemNumber}: Otros materiales redondeados y sumados a total. Total parcial otros materiales:`, otherMaterialsTotal);


        }); // End of itemBlocks.forEach

        console.log("Fin del procesamiento de ítems.");
         console.log("Acumuladores de paneles finales (fraccional/redondeado):", panelAccumulators);
        console.log("Errores totales encontrados:", currentErrorMessages);


        // --- Final Calculation of Panels from Accumulators ---
        let finalPanelTotals = {};
         for (const type in panelAccumulators) {
             if (panelAccumulators.hasOwnProperty(type)) {
                 const acc = panelAccumulators[type];
                 // Apply ceiling only to the fractional sum before adding to the rounded sum
                 const totalPanelsForType = roundUpFinalUnit(acc.suma_fraccionaria_pequenas) + acc.suma_redondeada_otros;

                 if (totalPanelsForType > 0) {
                      // Store final rounded panel total with descriptive name
                      finalPanelTotals[`Paneles de ${type}`] = totalPanelsForType;
                 }
             }
         }
         console.log("Totales finales de paneles (redondeo aplicado a fraccional + suma de redondeados):", finalPanelTotals);

        // --- Combine Final Panels with Other Materials Total ---
        let finalTotalMaterials = { ...finalPanelTotals, ...otherMaterialsTotal };
        console.log("Total final de materiales combinados (Paneles + Otros):", finalTotalMaterials);


        // --- Display Results ---
        if (currentErrorMessages.length > 0) {
            console.log("Mostrando mensajes de error.");
            // Display errors first if any
             resultsContent.innerHTML = '<div class="error-message"><h2>Errores de Validación:</h2>' +
                                        currentErrorMessages.map(msg => `<p>${msg}</p>`).join('') +
                                        '<p>Por favor, corrige los errores en los ítems marcados.</p></div>';
             // Clear/reset previous results and hide download buttons
             downloadOptionsDiv.classList.add('hidden');
             lastCalculatedTotalMaterials = {};
             lastCalculatedItemsSpecs = []; // Clear specs if there are errors in any item
             lastErrorMessages = currentErrorMessages; // Store errors for potential future handling
             return; // Stop here if there are errors
        }

        console.log("No se encontraron errores de validación. Generando resultados HTML.");
        // If no errors, proceed to display results and store them

        let resultsHtml = '<div class="report-header">';
        resultsHtml += '<h2>Resumen de Materiales</h2>';
        resultsHtml += `<p>Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}</p>`; // Format date for Spanish
        resultsHtml += '</div>';
        resultsHtml += '<hr>';

         // Display individual item summaries for valid items
        if (currentCalculatedItemsSpecs.length > 0) {
             console.log("Generando resumen de ítems calculados.");
             resultsHtml += '<h3>Detalle de Ítems Calculados:</h3>';
             currentCalculatedItemsSpecs.forEach(item => {
                 resultsHtml += `<div class="item-summary">`;
                 resultsHtml += `<h4>${getItemTypeName(item.type)} #${item.number}</h4>`;
                 resultsHtml += `<p><strong>Tipo:</strong> <span>${getItemTypeName(item.type)}</span></p>`;

                 if (item.type === 'muro') {
                     if (!isNaN(item.faces)) resultsHtml += `<p><strong>Nº Caras:</strong> <span>${item.faces}</span></p>`;
                     if (item.cara1PanelType) resultsHtml += `<p><strong>Panel Cara 1:</strong> <span>${item.cara1PanelType}</span></p>`;
                     if (item.faces === 2 && item.cara2PanelType) resultsHtml += `<p><strong>Panel Cara 2:</strong> <span>${item.cara2PanelType}</span></p>`;
                     if (!isNaN(item.postSpacing)) resultsHtml += `<p><strong>Espaciamiento Postes:</strong> <span>${item.postSpacing.toFixed(2)} m</span></p>`;
                     resultsHtml += `<p><strong>Estructura Doble:</strong> <span>${item.isDoubleStructure ? 'Sí' : 'No'}</span></p>`;

                     resultsHtml += `<p><strong>Segmentos:</strong></p>`;
                     if (item.segments && item.segments.length > 0) {
                         item.segments.forEach(seg => {
                            resultsHtml += `<p style="margin-left: 20px;">- Segmento ${seg.number}: ${seg.width.toFixed(2)} m (Ancho) x ${seg.height.toFixed(2)} m (Alto)</p>`;
                         });
                         if (!isNaN(item.totalMuroArea)) {
                              resultsHtml += `<p style="margin-left: 20px;"><strong>Área Total Segmentos:</strong> ${item.totalMuroArea.toFixed(2)} m²</p>`;
                         }
                         if (!isNaN(item.totalMuroWidth)) {
                             resultsHtml += `<p style="margin-left: 20px;"><strong>Ancho Total Segmentos:</strong> ${item.totalMuroWidth.toFixed(2)} m</p>`;
                         }
                     } else {
                          resultsHtml += `<p style="margin-left: 20px;">- Sin segmentos válidos</p>`;
                     }


                 } else if (item.type === 'cielo') {
                     if (item.cieloPanelType) resultsHtml += `<p><strong>Tipo de Panel:</strong> <span>${item.cieloPanelType}</span></p>`;
                     if (!isNaN(item.plenum)) resultsHtml += `<p><strong>Pleno:</strong> <span>${item.plenum.toFixed(2)} m</span></p>`;

                     resultsHtml += `<p><strong>Segmentos:</strong></p>`;
                     if (item.segments && item.segments.length > 0) {
                         item.segments.forEach(seg => {
                             resultsHtml += `<p style="margin-left: 20px;">- Segmento ${seg.number}: ${seg.width.toFixed(2)} m (Ancho) x ${seg.length.toFixed(2)} m (Largo)</p>`;
                         });
                         if (!isNaN(item.totalCieloArea)) {
                             resultsHtml += `<p style="margin-left: 20px;"><strong>Área Total Segmentos:</strong> ${item.totalCieloArea.toFixed(2)} m²</p>`;
                         }
                          if (!isNaN(item.totalCieloPerimeterSum)) {
                             resultsHtml += `<p style="margin-left: 20px;"><strong>Suma Anchos+Largos Segmentos:</strong> ${item.totalCieloPerimeterSum.toFixed(2)} m</p>`;
                         }
                     } else {
                         resultsHtml += `<p style="margin-left: 20px;">- Sin segmentos válidos</p>`;
                     }
                 }
                 resultsHtml += `</div>`;
             });
             resultsHtml += '<hr>';
        } else {
             console.log("No hay ítems calculados válidamente para mostrar resumen individual.");
        }


        resultsHtml += '<h3>Totales de Materiales (Cantidades a Comprar):</h3>';

        // finalTotalMaterials now holds the combined totals
        let hasMaterials = Object.keys(finalTotalMaterials).length > 0;
        console.log("Total final de materiales calculados (combinados):", finalTotalMaterials);

        if (hasMaterials) {
             console.log("Generando tabla de materiales totales.");
             // Sort materials alphabetically for consistent display
             const sortedMaterials = Object.keys(finalTotalMaterials).sort();

            resultsHtml += '<table><thead><tr><th>Material</th><th>Cantidad</th><th>Unidad</th></tr></thead><tbody>';

            sortedMaterials.forEach(material => {
                const cantidad = finalTotalMaterials[material];
                const unidad = getMaterialUnit(material); // Get unit using the helper function
                // Display the material name
                 resultsHtml += `<tr><td>${material}</td><td>${cantidad}</td><td>${unidad}</td></tr>`;
            });

            resultsHtml += '</tbody></table>';
            downloadOptionsDiv.classList.remove('hidden'); // Show download options
        } else {
             console.log("No se calcularon materiales totales positivos.");
             // If there are no materials after calculation but no validation errors, something went wrong or inputs resulted in 0.
             resultsHtml += '<p>No se pudieron calcular los materiales con las dimensiones ingresadas. Revisa los valores.</p>';
             downloadOptionsDiv.classList.add('hidden'); // Hide download options
        }

        // Append the generated HTML to the results content area
        resultsContent.innerHTML = resultsHtml; // Use = to replace previous content

        console.log("Resultados HTML generados y añadidos al DOM.");

        // Store the successfully calculated and rounded results and specs
        lastCalculatedTotalMaterials = finalTotalMaterials; // Store the final combined totals
        lastCalculatedItemsSpecs = currentCalculatedItemsSpecs; // Store specs of items included in calculation
        lastErrorMessages = []; // Clear errors as calculation was successful

         console.log("Estado de resultados almacenado para descarga.");
    };


    // --- PDF Generation Function ---
    const generatePDF = () => {
        console.log("Iniciando generación de PDF...");
       // Ensure there are calculated results to download
       if (Object.keys(lastCalculatedTotalMaterials).length === 0 || lastCalculatedItemsSpecs.length === 0) {
           console.warn("No hay resultados calculados para generar el PDF.");
           alert("Por favor, realiza un cálculo válido antes de generar el PDF.");
           return;
       }

       // Initialize jsPDF
       const { jsPDF } = window.jspdf;
       const doc = new jsPDF();

       // Define colors in RGB from CSS variables (using approximations based on common web colors)
       const primaryOliveRGB = [85, 107, 47]; // #556B2F
       const secondaryOliveRGB = [128, 128, 0]; // #808000
       const darkGrayRGB = [51, 51, 51]; // #333
       const mediumGrayRGB = [102, 102, 102]; // #666
       const lightGrayRGB = [224, 224, 224]; // #e0e0e0
       const extraLightGrayRGB = [248, 248, 248]; // #f8f8f8


       // --- Add Header ---
       doc.setFontSize(18);
       doc.setTextColor(primaryOliveRGB[0], primaryOliveRGB[1], primaryOliveRGB[2]);
       doc.setFont("helvetica", "bold"); // Use a standard font or include custom fonts
       doc.text("Resumen de Materiales Tablayeso", 14, 22);

       doc.setFontSize(10);
       doc.setTextColor(mediumGrayRGB[0], mediumGrayRGB[1], mediumGrayRGB[2]);
        doc.setFont("helvetica", "normal");
       doc.text(`Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}`, 14, 28);

       // Set starting Y position for the next content block
       let finalY = 35; // Start below the header

       // --- Add Item Summaries ---
       if (lastCalculatedItemsSpecs.length > 0) {
            console.log("Añadiendo resumen de ítems al PDF.");
            doc.setFontSize(14);
            doc.setTextColor(secondaryOliveRGB[0], secondaryOliveRGB[1], secondaryOliveRGB[2]);
           doc.setFont("helvetica", "bold");
            doc.text("Detalle de Ítems Calculados:", 14, finalY + 10);
            finalY += 15; // Move Y below the title

           const itemSummaryLineHeight = 5; // Space between summary lines within an item
           const itemBlockSpacing = 8; // Space between different item summaries

            lastCalculatedItemsSpecs.forEach(item => {
                // Add item title
                doc.setFontSize(10);
                doc.setTextColor(primaryOliveRGB[0], primaryOliveRGB[1], primaryOliveRGB[2]);
                doc.setFont("helvetica", "bold");
                doc.text(`${getItemTypeName(item.type)} #${item.number}:`, 14, finalY + itemSummaryLineHeight);
                finalY += itemSummaryLineHeight * 1.5; // Move down after the title

                // Add general item details (indented)
                doc.setFontSize(9);
                doc.setTextColor(darkGrayRGB[0], darkGrayRGB[1], darkGrayRGB[2]);
                doc.setFont("helvetica", "normal");

                doc.text(`Tipo: ${getItemTypeName(item.type)}`, 20, finalY + itemSummaryLineHeight);
                finalY += itemSummaryLineHeight;

                // Add type-specific details
                if (item.type === 'muro') {
                     if (!isNaN(item.faces)) {
                          doc.text(`Nº Caras: ${item.faces}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                     if (item.cara1PanelType) {
                          doc.text(`Panel Cara 1: ${item.cara1PanelType}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                     if (item.faces === 2 && item.cara2PanelType) {
                          doc.text(`Panel Cara 2: ${item.cara2PanelType}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                      if (!isNaN(item.postSpacing)) {
                         doc.text(`Espaciamiento Postes: ${item.postSpacing.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                      doc.text(`Estructura Doble: ${item.isDoubleStructure ? 'Sí' : 'No'}`, 20, finalY + itemSummaryLineHeight);
                       finalY += itemSummaryLineHeight;

                     doc.text(`Segmentos:`, 20, finalY + itemSummaryLineHeight);
                     finalY += itemSummaryLineHeight;
                      if (item.segments && item.segments.length > 0) {
                         item.segments.forEach(seg => {
                            doc.text(`- Segmento ${seg.number}: ${seg.width.toFixed(2)} m (Ancho) x ${seg.height.toFixed(2)} m (Alto)`, 25, finalY + itemSummaryLineHeight);
                            finalY += itemSummaryLineHeight;
                         });
                          if (!isNaN(item.totalMuroArea)) {
                              doc.text(`- Área Total Segmentos: ${item.totalMuroArea.toFixed(2)} m²`, 25, finalY + itemSummaryLineHeight);
                              finalY += itemSummaryLineHeight;
                          }
                           if (!isNaN(item.totalMuroWidth)) {
                              doc.text(`- Ancho Total Segmentos: ${item.totalMuroWidth.toFixed(2)} m`, 25, finalY + itemSummaryLineHeight);
                              finalY += itemSummaryLineHeight;
                          }

                      } else {
                           doc.text(`- Sin segmentos válidos`, 25, finalY + itemSummaryLineHeight);
                           finalY += itemSummaryLineHeight;
                      }


                } else if (item.type === 'cielo') {
                     if (item.cieloPanelType) {
                          doc.text(`Tipo de Panel: ${item.cieloPanelType}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                    if (!isNaN(item.plenum)) {
                        doc.text(`Pleno: ${item.plenum.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight);
                         finalY += itemSummaryLineHeight;
                    }

                     doc.text(`Segmentos:`, 20, finalY + itemSummaryLineHeight);
                     finalY += itemSummaryLineHeight;
                     if (item.segments && item.segments.length > 0) {
                         item.segments.forEach(seg => {
                             doc.text(`- Segmento ${seg.number}: ${seg.width.toFixed(2)} m (Ancho) x ${seg.length.toFixed(2)} m (Largo)`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                         });
                         if (!isNaN(item.totalCieloArea)) {
                             doc.text(`- Área Total Segmentos: ${item.totalCieloArea.toFixed(2)} m²`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                         }
                          if (!isNaN(item.totalCieloPerimeterSum)) {
                             doc.text(`- Suma Anchos+Largos Segmentos: ${item.totalCieloPerimeterSum.toFixed(2)} m`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                         }
                     } else {
                         doc.text(`- Sin segmentos válidos`, 25, finalY + itemSummaryLineHeight);
                         finalY += itemSummaryLineHeight;
                     }
                }
                finalY += itemBlockSpacing; // Add space after each item summary block
            });
            finalY += 5; // Add space before the total materials table title
       } else {
            console.log("No hay ítems calculados válidamente para añadir resumen al PDF.");
       }


       // --- Add Total Materials Table ---
        console.log("Añadiendo tabla de materiales totales al PDF.");
       doc.setFontSize(14);
       doc.setTextColor(secondaryOliveRGB[0], secondaryOliveRGB[1], secondaryOliveRGB[2]);
       doc.setFont("helvetica", "bold");
       doc.text("Totales de Materiales:", 14, finalY + 10);
       finalY += 15; // Move Y below the title

       const tableColumn = ["Material", "Cantidad", "Unidad"];
       const tableRows = [];

       // Prepare data for the table
       const sortedMaterials = Object.keys(lastCalculatedTotalMaterials).sort();
       sortedMaterials.forEach(material => {
           const cantidad = lastCalculatedTotalMaterials[material];
           const unidad = getMaterialUnit(material); // Get unit using the helper function
           // Use the material name directly from the key
            tableRows.push([material, cantidad, unidad]);
       });

        // Add the table using jspdf-autotable
        doc.autoTable({
            head: [tableColumn],
            body: tableRows,
            startY: finalY, // Start position below the last content
            theme: 'plain', // Start with a plain theme to apply custom styles
            headStyles: {
                fillColor: lightGrayRGB,
                textColor: darkGrayRGB,
                fontStyle: 'bold',
                halign: 'center', // Horizontal alignment
                valign: 'middle', // Vertical alignment
                lineWidth: 0.1,
                lineColor: lightGrayRGB,
                fontSize: 10 // Match HTML table header font size
            },
            bodyStyles: {
                textColor: darkGrayRGB,
                lineWidth: 0.1,
                lineColor: lightGrayRGB,
                fontSize: 9 // Match HTML table body font size
            },
             alternateRowStyles: { // Styling for alternate rows
                fillColor: extraLightGrayRGB,
            },
             // Specific column styles (Cantidad column is the second one, index 1)
            columnStyles: {
                1: {
                    halign: 'right', // Align quantity to the right
                    fontStyle: 'bold',
                    textColor: primaryOliveRGB // For quantity text color
                },
                 2: { // Unit column
                    halign: 'center' // Align unit to the center or left as preferred
                }
            },
            margin: { top: 10, right: 14, bottom: 14, left: 14 }, // Add margin
             didDrawPage: function (data) {
               // Optional: Add page number or footer here
               doc.setFontSize(8);
               doc.setTextColor(mediumGrayRGB[0], mediumGrayRGB[1], mediumGrayRGB[2]);
               // Using `doc.internal.pageSize.getWidth()` to center the footer roughly
               const footerText = '© 2025 [PROPUL] - Calculadora de Materiales Tablayeso v2.0'; // Replaced placeholder
               const textWidth = doc.getStringUnitWidth(footerText) * doc.internal.getFontSize() / doc.internal.scaleFactor;
               const centerX = (doc.internal.pageSize.getWidth() - textWidth) / 2;
               doc.text(footerText, centerX, doc.internal.pageSize.height - 10);

               // Add simple page number
               const pageNumberText = `Página ${data.pageNumber}`;
               const pageNumberWidth = doc.getStringUnitWidth(pageNumberText) * doc.internal.getFontSize() / doc.internal.scaleFactor;
               const pageNumberX = doc.internal.pageSize.getWidth() - data.settings.margin.right - pageNumberWidth;
               doc.text(pageNumberText, pageNumberX, doc.internal.pageSize.height - 10);
            }
        });

        // Update finalY after the table
       finalY = doc.autoTable.previous.finalY;

       console.log("PDF generado.");
       // --- Save the PDF ---
       doc.save(`Calculo_Materiales_${new Date().toLocaleDateString('es-ES').replace(/\//g, '-')}.pdf`); // Filename with date
   };

// --- Excel Generation Function ---
const generateExcel = () => {
    console.log("Iniciando generación de Excel...");
    // Ensure there are calculated results to download
   if (Object.keys(lastCalculatedTotalMaterials).length === 0 || lastCalculatedItemsSpecs.length === 0) {
       console.warn("No hay resultados calculados para generar el Excel.");
       alert("Por favor, realiza un cálculo válido antes de generar el Excel.");
       return;
   }

    // Assumes you have loaded the xlsx library
   if (typeof XLSX === 'undefined') {
        console.error("La librería xlsx no está cargada.");
        alert("Error al generar Excel: Librería xlsx no encontrada.");
        return;
   }


   // Data for the Excel sheet
   let sheetData = [];

   // Add Header
   sheetData.push(["Calculadora de Materiales Tablayeso"]);
   sheetData.push([`Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}`]);
   sheetData.push([]); // Blank row for spacing

   // Add Item Summaries
    console.log("Añadiendo resumen de ítems al Excel.");
   sheetData.push(["Detalle de Ítems Calculados:"]);
    // --- MODIFIED HEADERS: Added explicit Total Width and Total Length columns ---
   sheetData.push([
       "Tipo Item", "Número Item", "Detalle/Dimensiones",
       "Nº Caras", "Panel Cara 1", "Panel Cara 2", "Tipo Panel Cielo",
       "Espaciamiento Postes (m)", "Pleno (m)", "Estructura Doble",
       "Ancho Total (m)", "Largo Total (m)", "Área Total (m²)" // New columns
   ]);
   // --- End MODIFIED HEADERS ---


   lastCalculatedItemsSpecs.forEach(item => {
        if (item.type === 'muro') {
            // --- Details common to this muro item (config only) ---
            const muroCommonDetails = [
                getItemTypeName(item.type), // 0: Tipo Item
                item.number,                 // 1: Número Item
                '',                          // 2: Placeholder for Detail/Dimensiones
                !isNaN(item.faces) ? item.faces : '', // 3: Nº Caras
                item.cara1PanelType ? item.cara1PanelType : '', // 4: Panel Cara 1
                item.faces === 2 && item.cara2PanelType ? item.cara2PanelType : '', // 5: Panel Cara 2
                '',                          // 6: Tipo Panel Cielo (empty for muro)
                !isNaN(item.postSpacing) ? item.postSpacing.toFixed(2) : '', // 7: Espaciamiento Postes
                '',                          // 8: Pleno (empty for muro)
                item.isDoubleStructure ? 'Sí' : 'No' // 9: Estructura Doble
                // Note: Total columns (10, 11, 12) are NOT in commonDetails array
            ];

            // Add the main row for item options and totals
            const muroSummaryRow = [...muroCommonDetails]; // Copy common details
            muroSummaryRow[2] = 'Opciones:'; // Set detail column for this row
            // --- Populate Total Columns for Summary Row ---
            muroSummaryRow.push(          // Push to add to the end of the copied common details
               !isNaN(item.totalMuroWidth) ? item.totalMuroWidth.toFixed(2) : '', // 10: Ancho Total (m)
               '',                                                               // 11: Largo Total (m) - Blank for muro
               !isNaN(item.totalMuroArea) ? item.totalMuroArea.toFixed(2) : ''     // 12: Área Total (m²)
            );
            // --- End Populate Total Columns ---
            sheetData.push(muroSummaryRow);

            // Add segments label row (repeat common details, just change the detail column)
             const muroSegmentsLabelRow = [...muroCommonDetails];
             muroSegmentsLabelRow[2] = 'Segmentos:'; // Set detail column for this row
             // Add empty cells for the total columns
             muroSegmentsLabelRow.push('', '', ''); // 10, 11, 12
             sheetData.push(muroSegmentsLabelRow);


             if (item.segments && item.segments.length > 0) {
                 item.segments.forEach(seg => {
                      // Create a row for each segment, repeating common item details
                      const segmentRow = [...muroCommonDetails]; // Copy common details for this row
                      segmentRow[2] = `- Seg ${seg.number}: ${seg.width.toFixed(2)}m x ${seg.height.toFixed(2)}m`; // Set segment dimensions in Detail column
                      // Add empty cells for the total columns (repeat blank totals for each segment row)
                      segmentRow.push('', '', ''); // 10, 11, 12
                      sheetData.push(segmentRow);
                 });
              } else {
                   // Row for "Sin segmentos válidos" (repeat common details)
                   const noSegmentsRow = [...muroCommonDetails];
                   noSegmentsRow[2] = `- Sin segmentos válidos`; // Set detail column
                    // Add empty cells for the total columns
                    noSegmentsRow.push('', '', ''); // 10, 11, 12
                    sheetData.push(noSegmentsRow);
              }


        } else if (item.type === 'cielo') {
            // --- Details common to this cielo item (config only) ---
            const cieloCommonDetails = [
                getItemTypeName(item.type), // 0: Tipo Item
                item.number,                 // 1: Número Item
                '',                          // 2: Placeholder for Detail/Dimensiones
                '',                          // 3: Nº Caras (empty for cielo)
                '', '',                     // 4, 5: Panel Cara 1, Cara 2 (empty for cielo)
                item.cieloPanelType ? item.cieloPanelType : '', // 6: Tipo Panel Cielo
                '',                          // 7: Espaciamiento Postes (empty for cielo)
                !isNaN(item.plenum) ? item.plenum.toFixed(2) : '', // 8: Pleno
                ''                           // 9: Estructura Doble (empty for cielo)
                 // Note: Total columns (10, 11, 12) are NOT in commonDetails array
            ];

             // Add the main row for item options and totals
            const cieloSummaryRow = [...cieloCommonDetails]; // Copy common details
            cieloSummaryRow[2] = 'Opciones:'; // Set detail column
            // --- Populate Total Columns for Summary Row ---
            cieloSummaryRow.push(          // Push to add to the end
               '',                                                              // 10: Ancho Total (m) - Blank for cielo
               '',                                                              // 11: Largo Total (m) - Blank for cielo
               !isNaN(item.totalCieloArea) ? item.totalCieloArea.toFixed(2) : '' // 12: Área Total (m²)
            );
            // --- End Populate Total Columns ---
            sheetData.push(cieloSummaryRow);

            // Add segments label row (repeat common details, just change the detail column)
             const cieloSegmentsLabelRow = [...cieloCommonDetails];
             cieloSegmentsLabelRow[2] = 'Segmentos:'; // Set detail column for this row
              // Add empty cells for the total columns
              cieloSegmentsLabelRow.push('', '', ''); // 10, 11, 12
              sheetData.push(cieloSegmentsLabelRow);


              if (item.segments && item.segments.length > 0) {
                  item.segments.forEach(seg => {
                      // Create a row for each segment, repeating common item details
                      const segmentRow = [...cieloCommonDetails]; // Copy common details for this row
                      segmentRow[2] = `Seg ${seg.number}: ${seg.width.toFixed(2)}m x ${seg.length.toFixed(2)}m`; // Set segment dimensions
                       // Add empty cells for the total columns (repeat blank totals for each segment row)
                       segmentRow.push('', '', ''); // 10, 11, 12
                       sheetData.push(segmentRow);
                  });
               } else {
                   // Row for "Sin segmentos válidos" (repeat common details)
                   const noSegmentsRow = [...cieloCommonDetails];
                   noSegmentsRow[2] = `- Sin segmentos válidos`; // Set detail column
                    // Add empty cells for the total columns
                    noSegmentsRow.push('', '', ''); // 10, 11, 12
                    sheetData.push(noSegmentsRow);
               }
        }
   });
    sheetData.push([]); // Blank row for spacing

   // Add Total Materials Table
    console.log("Añadiendo tabla de materiales totales al Excel.");
   sheetData.push(["Totales de Materiales (Cantidades a Comprar):"]);
   sheetData.push(["Material", "Cantidad", "Unidad"]);

   const sortedMaterials = Object.keys(lastCalculatedTotalMaterials).sort();
   sortedMaterials.forEach(material => {
       const cantidad = lastCalculatedTotalMaterials[material];
       const unidad = getMaterialUnit(material);
       // Use the material name directly from the key
        sheetData.push([material, cantidad, unidad]);
   });

   // Create a workbook and worksheet
   const wb = XLSX.utils.book_new();
   const ws = XLSX.utils.aoa_to_sheet(sheetData);

   // Optional: Add some basic styling (e.g., bold headers)
   // Note: XLSX.js basic functionality focuses on data, styling is limited
   // More advanced styling requires libraries like exceljs or server-side generation.
   // For basic bolding, you might manually format cells if needed, but auto-sheet is simpler.
   // For headers, you might iterate through the cell references like A1, B1 etc.

   // Add the worksheet to the workbook
   XLSX.utils.book_append_sheet(wb, ws, "CalculoMateriales");

   // Generate and save the Excel file
   XLSX.writeFile(wb, `Calculo_Materiales_${new Date().toLocaleDateString('es-ES').replace(/\//g, '-')}.xlsx`); // Filename with date
   console.log("Excel generado.");
};

// ... (rest of the code remains the same)
// --- Event Listeners ---
addItemBtn.addEventListener('click', createItemBlock);
calculateBtn.addEventListener('click', calculateMaterials);
generatePdfBtn.addEventListener('click', generatePDF);
generateExcelBtn.addEventListener('click', generateExcel);

// --- Initial Setup ---
createItemBlock(); // Add one item by default on page load
toggleCalculateButtonState(); // Set initial state of calculate button


}); // End DOMContentLoaded
