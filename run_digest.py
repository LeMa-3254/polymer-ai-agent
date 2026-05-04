#!/usr/bin/env python3
"""
Weekly Polymer AI Digest generator — 2026-05-04
Synthesizes research from 7 web searches into Word + PowerPoint.
"""

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from agent import generate_word, generate_pptx

# ── Synthesized research data ──────────────────────────────────────────────────

data = {
    "date": "2026-05-04",

    "headline": (
        "From lab-validated generative design to 60 % recycling gains: "
        "AI is reshaping every stage of the polymer lifecycle."
    ),

    "executive_summary": (
        "The spring 2026 landscape for AI in polymer materials science is defined by a "
        "decisive shift from prediction to creation. Georgia Tech's March 2026 debut of "
        "POLYT5 and polyBART — chemical-language models that generate 100 % chemically "
        "valid polymer structures from target property specifications — marks the first "
        "lab-validated generative AI system for polymer design, published in Nature npj "
        "Artificial Intelligence. Simultaneously, multimodal large language models such "
        "as PolyLLMem (Llama 3 + Uni-Mol) and MatterChat are pushing polymer property "
        "prediction into low-data regimes where traditional QSPR models struggle.\n\n"

        "On the characterisation front, 2025 proved to be a turning point for AI-assisted "
        "vibrational spectroscopy. Deep learning models now reconstruct low-quality FTIR "
        "and Raman spectra, predict polymer fingerprint regions with up to 87 % accuracy, "
        "and classify microplastics with 99 % accuracy using multi-modal spectral fusion. "
        "These advances are enabling rapid quality control and environmental monitoring "
        "at industrial scale without the need for expert spectroscopists.\n\n"

        "Process optimisation is maturing from proof-of-concept into production deployment. "
        "CatBoost-based surrogate models paired with genetic algorithms have reduced "
        "injection-moulding cycle times by 4.5 % while maintaining part quality. Closed-loop "
        "ML controllers for recycled-polypropylene blends address feedstock variability — a "
        "critical challenge as sustainability mandates drive greater use of post-consumer "
        "resin. The SPE now offers dedicated workshops and on-demand training on AI-driven "
        "injection moulding, signalling mainstream industry adoption.\n\n"

        "The sustainability dimension is perhaps the most commercially consequential: "
        "AI-enabled sorting systems report 60 % efficiency gains and up to 95 % separation "
        "accuracy. Nature Reviews Materials published a landmark review in 2026 cataloguing "
        "AI as a primary driver of sustainable materials and circularity. New benchmark "
        "databases — POINT2, OpenPoly (3 985 data points, 26 properties), and the "
        "PolyMetriX Python ecosystem — are standardising the evaluation of ML models "
        "and reducing the data-scarcity bottleneck that has historically constrained "
        "polymer informatics."
    ),

    "findings": [
        {
            "category": "Generative Design",
            "title": "POLYT5 & polyBART: First Lab-Validated Generative AI for Polymer Design",
            "description": (
                "Georgia Tech researchers introduced POLYT5 and polyBART in March 2026 — "
                "chemical-language transformer models that generate new polymer SMILES "
                "representations directly from desired property targets, guaranteeing "
                "100 % chemical validity. The team synthesised and characterised a novel "
                "dielectric material designed entirely by the model, validating the "
                "approach in a physical laboratory setting. The work was published in "
                "Nature npj Artificial Intelligence and represents the first end-to-end "
                "lab-confirmed generative polymer design pipeline."
            ),
            "source": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "PolyLLMem & MatterChat: Multimodal LLMs Enter Polymer Property Prediction",
            "description": (
                "PolyLLMem integrates Llama 3 text embeddings with Uni-Mol molecular "
                "structure embeddings to predict multiple polymer properties from a "
                "unified multimodal representation, excelling in low-data regimes where "
                "traditional QSPR models fail. In parallel, MatterChat — published in "
                "Nature Machine Intelligence in 2026 — provides interpretable reasoning "
                "alongside high-precision property predictions, bridging the gap between "
                "LLM usability and scientific rigour. A June 2025 benchmarking study "
                "confirmed that fine-tuned LLaMA-3 consistently outperforms GPT-3.5 for "
                "polymer tasks, underscoring the value of open-source fine-tuning."
            ),
            "source": "https://pubs.acs.org/doi/10.1021/acs.chemmater.5c00940",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "Graph Neural Networks and CNN-LSTM Hybrids Push Polymer ML Accuracy",
            "description": (
                "A PubMed-indexed 2025 review catalogues GNN architectures purpose-built "
                "for polymer characterisation and property prediction, noting that message-"
                "passing networks capture repeat-unit topology better than fingerprint-based "
                "models. Complementary work using Wasserstein GAN-based data augmentation "
                "combined with CNN-LSTM architectures achieved R² = 0.95 on mechanical "
                "property benchmarks — approaching the accuracy ceiling of experimental "
                "measurement uncertainty. A hybrid framework leveraging both composition "
                "and sequence features further improved generalisation across polymer "
                "families (PubMed, July 2025)."
            ),
            "source": "https://pubmed.ncbi.nlm.nih.gov/41607004/",
            "impact": "high",
        },
        {
            "category": "Characterization",
            "title": "AI Spectroscopy: 99 % Microplastic ID and Deep-Learning FTIR Reconstruction",
            "description": (
                "A multi-modal 1D-CNN trained on Raman and ATR-FTIR data with three-level "
                "data fusion achieved 99 % microplastic classification accuracy, reported "
                "in PMC 2025. A separate MRS Advances (2026) study used an MLP to predict "
                "polymer FTIR fingerprint regions, reaching 87 % accuracy for PS, 84 % for "
                "HDPE, and 81 % for PET — enabling in-silico screening without running "
                "physical spectra. Generative AI models are also being used to synthesise "
                "training spectra, addressing the chronic scarcity of labelled spectroscopic "
                "datasets for novel or degraded polymers."
            ),
            "source": "https://pmc.ncbi.nlm.nih.gov/articles/PMC12044591/",
            "impact": "high",
        },
        {
            "category": "Processing",
            "title": "CatBoost + Genetic Algorithms Cut Injection-Moulding Cycle Time by 4.5 %",
            "description": (
                "An integrated framework pairing a CatBoost surrogate model with a genetic "
                "algorithm optimiser demonstrated a 4.5 % reduction in injection-moulding "
                "cycle time while maintaining dimensional and mechanical part quality, "
                "published in MDPI AI (2025). A complementary PMC study applied ANNs and "
                "polynomial regression to closed-loop process control for recycled-PP blends, "
                "adaptively adjusting barrel temperatures and injection speeds to compensate "
                "for melt-flow variability in post-consumer resin. The SPE formalised "
                "industry training with dedicated AI-for-injection-moulding workshops in 2025."
            ),
            "source": "https://pmc.ncbi.nlm.nih.gov/articles/PMC11991421/",
            "impact": "medium",
        },
        {
            "category": "Sustainability",
            "title": "AI-Enabled Sorting Achieves 60 % Efficiency Gains and 95 % Separation Accuracy",
            "description": (
                "Plastics Today (2026) reported AI-driven recycling systems delivering a "
                "60 % efficiency increase, with ML classifiers applied to hyperspectral "
                "imaging, FTIR, Raman, and LIBS sensor data reaching up to 95 % material-"
                "separation accuracy. Nature Reviews Materials published a 2026 landmark "
                "review positioning AI as a primary enabler of sustainable materials and "
                "the circular economy. AI-assisted kinetic modelling and enzyme design are "
                "also accelerating chemical recycling routes — pyrolysis, solvolysis, and "
                "biocatalytic depolymerisation — towards commercial viability."
            ),
            "source": "https://www.plasticstoday.com/packaging/5-game-changing-recycling-predictions-for-2026",
            "impact": "high",
        },
        {
            "category": "Generative Design",
            "title": "polyGT Decoder-Only Transformer Enables Comprehensive Polymer Inverse Design",
            "description": (
                "The polyGT model, published in Chemical Engineering Science (2025), is a "
                "GPT-style decoder-only transformer trained on a large polymer SMILES corpus "
                "that supports both property comprehension and target-driven inverse design "
                "in a single architecture. Unlike encoder-only models such as polyBERT, "
                "polyGT can generate novel repeat-unit sequences conditioned on multi-"
                "property targets, making it suitable for simultaneous optimisation of "
                "dielectric, thermal, and mechanical specifications. The model represents "
                "a step towards a unified foundation model for polymer informatics."
            ),
            "source": "https://www.sciencedirect.com/science/article/abs/pii/S0009250925020159",
            "impact": "medium",
        },
        {
            "category": "Sustainability",
            "title": "AI Guides Inverse Design of Recyclable Vitrimeric Polymers",
            "description": (
                "A 2025 Advanced Science paper combined molecular-dynamics simulation with "
                "ML surrogate models to inversely design vitrimers with targeted glass-"
                "transition temperatures. Synthesised candidates demonstrated both "
                "healability and high-temperature flowability — key properties for "
                "closed-loop mechanical recycling of thermoset-like materials. This "
                "represents a rare case where AI-driven inverse design directly addressed "
                "recyclability as a primary design objective rather than a post-hoc "
                "consideration."
            ),
            "source": "https://advanced.onlinelibrary.wiley.com/doi/10.1002/advs.202411385",
            "impact": "medium",
        },
        {
            "category": "LLMs",
            "title": "Fine-Tuned GPT-3.5 Predicts Polymer Solubility Across 3 373 Systems",
            "description": (
                "Researchers fine-tuned GPT-3.5 on a dataset of 3 373 polymers and 51 "
                "solvents to predict solubility, demonstrating that LLMs can serve as "
                "flexible drop-in replacements for specialised QSPR models, published "
                "in ACS Materials Letters (2025). The model is particularly useful for "
                "solvent-screening tasks where experimental data is sparse and traditional "
                "group-contribution methods struggle. The study also highlighted that "
                "prompt engineering and chain-of-thought reasoning further boost "
                "prediction accuracy without additional fine-tuning."
            ),
            "source": "https://pubs.acs.org/doi/full/10.1021/acsmaterialslett.5c00054",
            "impact": "medium",
        },
        {
            "category": "Property Prediction",
            "title": "POINT2 and OpenPoly: New Benchmark Databases Address Data-Scarcity Bottleneck",
            "description": (
                "POINT2 (POlymer INformatics Training and Testing, March 2025) provides a "
                "standardised benchmark protocol covering diverse property tasks, enabling "
                "rigorous model comparison across the community for the first time. OpenPoly "
                "contributes 3 985 experimentally validated polymer-property data points "
                "spanning 26 key properties derived from literature mining, published in "
                "Chinese Journal of Polymer Science (2025). Together these resources "
                "lower the barrier to entry for new ML practitioners and support "
                "reproducible benchmarking — a prerequisite for scientific credibility "
                "in the field."
            ),
            "source": "https://arxiv.org/html/2503.23491v1",
            "impact": "medium",
        },
    ],

    "key_trends": [
        "Generative AI transitions from property prediction to lab-validated inverse design of new polymers",
        "Multimodal LLMs (text + molecular structure) outperform single-modality models in low-data regimes",
        "AI-driven recycling sorting systems achieve >95 % accuracy, making circular-economy targets achievable",
        "Open benchmark databases (POINT2, OpenPoly) standardise model evaluation and accelerate reproducible science",
        "Closed-loop ML process control for recycled feedstocks becomes a production-readiness requirement",
    ],

    "notable_tools": [
        {
            "name": "POLYT5 / polyBART",
            "purpose": (
                "Georgia Tech chemical-language transformers that generate chemically valid "
                "polymer SMILES from property targets — the first lab-validated generative "
                "polymer design AI."
            ),
            "url": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
        },
        {
            "name": "PolyMetriX",
            "purpose": (
                "Open-source Python ecosystem (npj Computational Materials, 2025) providing "
                "curated polymer datasets, novel hierarchical featurisation, and end-to-end "
                "informatics workflows."
            ),
            "url": "https://www.nature.com/articles/s41524-025-01823-y",
        },
        {
            "name": "polyBERT",
            "purpose": (
                "Polymer chemical-language BERT model (Nature Communications) that generates "
                "dense fingerprints for ML property models — backbone of several 2025-2026 "
                "multimodal architectures."
            ),
            "url": "https://www.nature.com/articles/s41467-023-39868-6",
        },
        {
            "name": "MatterChat",
            "purpose": (
                "Multimodal LLM framework (Nature Machine Intelligence, 2026) integrating "
                "structural data with language models to deliver high-precision property "
                "predictions with interpretable reasoning."
            ),
            "url": "https://www.nature.com/articles/s42256-026-01214-y",
        },
        {
            "name": "OpenPoly Database",
            "purpose": (
                "Curated open database of 3 985 experimentally validated polymer-property "
                "data points across 26 properties, enabling benchmarking and multi-property "
                "ML model training."
            ),
            "url": "https://link.springer.com/article/10.1007/s10118-025-3402-y",
        },
        {
            "name": "POINT2 Benchmark",
            "purpose": (
                "Standardised polymer informatics training-and-testing protocol (arXiv, "
                "March 2025) providing reproducible benchmarks across diverse property "
                "prediction tasks for fair model comparison."
            ),
            "url": "https://arxiv.org/abs/2503.23491",
        },
    ],

    "outlook": (
        "The convergence of generative AI, multimodal LLMs, and robust open databases "
        "is set to compress the polymer R&D cycle from years to weeks. As POLYT5-class "
        "models mature, we anticipate tight integration with high-throughput synthesis "
        "platforms, creating autonomous design-make-test loops for functional polymers. "
        "Sustainability will remain the dominant commercial driver: regulatory pressure "
        "across the EU and US is accelerating adoption of AI-powered sorting and "
        "chemical-recycling optimisation, with the AI Circular Economy Conference (Cologne, "
        "March 2026) underscoring that circularity is now an operational imperative rather "
        "than a marketing claim. Benchmark standardisation through POINT2 and OpenPoly "
        "will improve reproducibility and draw greater investment from chemical majors. "
        "The remaining frontier — combining molecular-scale generative design with "
        "meso-scale processing simulation under a single differentiable pipeline — is "
        "likely to define the next wave of breakthrough publications through 2027."
    ),
}

# ── Generate documents ─────────────────────────────────────────────────────────

output_dir = Path(__file__).parent / "output"
output_dir.mkdir(exist_ok=True)

date_str = data["date"]
docx_path = output_dir / f"Polymer_AI_Digest_{date_str}.docx"
pptx_path = output_dir / f"Polymer_AI_Digest_{date_str}.pptx"

print("Generating Word document …")
generate_word(data, docx_path)

print("Generating PowerPoint …")
generate_pptx(data, pptx_path)

print(f"\nDone!  Files written to {output_dir.resolve()}")
