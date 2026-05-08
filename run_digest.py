#!/usr/bin/env python3
"""
Weekly Polymer AI Digest generator — 2026-05-08
Synthesises research from 9 web searches into Word + PowerPoint.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from agent import generate_word, generate_pptx

# ── Synthesised research data (9 searches, May 2026) ──────────────────────────

data = {
    "date": "2026-05-08",

    "headline": (
        "From lab-validated generative design to 60 % recycling efficiency gains: "
        "AI is reshaping every stage of the polymer lifecycle in 2026."
    ),

    "executive_summary": (
        "The spring 2026 landscape for AI in polymer materials science is defined by a "
        "decisive shift from prediction to creation. Georgia Tech's March 2026 debut of "
        "POLYT5 — a chemical-language transformer model that generates 100 % chemically "
        "valid polymer structures from target property specifications — marks the first "
        "lab-validated generative AI system for polymer design. The team synthesised a "
        "new dielectric material designed entirely by the model, validating the approach "
        "experimentally and publishing results in Nature npj Artificial Intelligence. "
        "In parallel, the polyGT decoder-only transformer (Chemical Engineering Science, "
        "2025) enables simultaneous multi-property inverse design, and a benchmarking "
        "study found CharRNN, REINVENT, and GraphINVENT to be the leading generative "
        "architectures for realistic polymer library exploration.\n\n"

        "Multimodal large language models are rapidly entering polymer property prediction. "
        "PolyLLMem fuses Llama 3 text embeddings with Uni-Mol molecular representations, "
        "while MatterChat (Nature Machine Intelligence, 2026) pairs high-precision "
        "property prediction with interpretable chain-of-thought reasoning. Fine-tuned "
        "GPT-3.5 now achieves 90 % accuracy on polymer solubility classification across "
        "3 373 solvent–polymer pairs, underscoring that LLMs are viable drop-in "
        "alternatives to bespoke QSPR models in low-data regimes. A June 2025 "
        "benchmarking study confirmed that fine-tuned LLaMA-3 consistently outperforms "
        "GPT-3.5 for polymer-specific tasks.\n\n"

        "Characterisation via AI-augmented spectroscopy reached new performance ceilings: "
        "a multi-modal 1D-CNN trained on fused Raman and ATR-FTIR data achieved 99 % "
        "microplastic classification accuracy, while an MRS Advances (2026) deep-learning "
        "model predicts FTIR fingerprint regions at 87 % accuracy for PS, enabling "
        "in-silico screening without physical spectra. On the processing side, CatBoost "
        "surrogate models paired with genetic algorithms reduced injection-moulding cycle "
        "times by 4.5 %, and physics-informed neural networks cut composite cure cycle "
        "times by up to 30 %. The SPE now offers dedicated AI-for-injection-moulding "
        "workshops, signalling mainstream industrial adoption.\n\n"

        "Sustainability is the most commercially consequential driver: AI-enabled sorting "
        "systems report 60 % efficiency gains and up to 95 % material-separation accuracy. "
        "Nature Reviews Materials published a landmark 2026 review positioning AI as a "
        "primary enabler of circular materials economies. New open resources — the "
        "OpenPoly database (3 985 data points, 26 properties) and the POINT2 benchmark "
        "protocol — are standardising model evaluation and lowering the barrier to entry "
        "for industrial polymer informatics practitioners."
    ),

    "findings": [
        {
            "category": "Generative Design",
            "title": "POLYT5: First Lab-Validated Generative AI for Polymer Design",
            "description": (
                "Georgia Tech researchers introduced POLYT5 in March 2026 — a "
                "chemical-language transformer model trained on the 'grammar' of polymer "
                "chemistry that generates new SMILES representations from desired property "
                "targets, guaranteeing 100 % chemical validity. The team synthesised and "
                "characterised a novel dielectric material designed entirely by the model, "
                "providing the first end-to-end laboratory confirmation of a generative "
                "polymer design pipeline. Published in Nature npj Artificial Intelligence, "
                "POLYT5 is openly accessible and already attracting commercial licensing "
                "interest from specialty-chemical manufacturers."
            ),
            "source": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "PolyLLMem & MatterChat: Multimodal LLMs Deliver Interpretable Polymer Predictions",
            "description": (
                "PolyLLMem (Chemistry of Materials, 2025) integrates Llama 3 text embeddings "
                "with Uni-Mol molecular structure embeddings via LoRA fine-tuning, "
                "excelling in low-data polymer property prediction tasks. MatterChat, "
                "published in Nature Machine Intelligence (2026), further adds interpretable "
                "reasoning to high-precision predictions, making AI decisions auditable for "
                "regulatory submissions. A June 2025 LLM benchmark study confirmed that "
                "fine-tuned open-source LLaMA-3 consistently outperforms GPT-3.5 for "
                "polymer tasks, underscoring the value of domain-specific fine-tuning over "
                "general-purpose closed models."
            ),
            "source": "https://pubs.acs.org/doi/10.1021/acs.chemmater.5c00940",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "CNN-LSTM + WGAN Hybrids Reach R² = 0.95 on Polymer Mechanical Benchmarks",
            "description": (
                "Combining Wasserstein GAN-based data augmentation with CNN-LSTM sequence "
                "models, researchers achieved R² = 0.95 and RMSE = 0.23 on polymer "
                "mechanical property benchmarks — approaching the accuracy ceiling set by "
                "experimental measurement uncertainty. A hybrid framework (PubMed, July "
                "2025) leveraging both polymer composition and sequence features further "
                "improved generalisation across diverse polymer families. A comprehensive "
                "2025 uncertainty-quantification study benchmarked nine UQ methods for "
                "key properties including Tg, band gap, melting temperature, and "
                "decomposition temperature, providing practitioners with reliability "
                "estimates essential for high-stakes design decisions."
            ),
            "source": "https://pubmed.ncbi.nlm.nih.gov/40619673/",
            "impact": "high",
        },
        {
            "category": "Characterization",
            "title": "AI Spectroscopy: 99 % Microplastic Classification and Deep-Learning FTIR Prediction",
            "description": (
                "A three-level fusion 1D-CNN combining Raman and ATR-FTIR data achieved "
                "99 % microplastic classification accuracy (PMC, 2025), enabling rapid "
                "environmental monitoring without expert spectroscopists. Separately, an "
                "MLP model published in MRS Advances (2026) predicts polymer FTIR "
                "fingerprint regions (1500–400 cm⁻¹) with 87 % accuracy for PS, "
                "84 % for HDPE, and 81 % for PET, allowing in-silico screening of "
                "candidate formulations. The open-source xpectrass toolkit standardises "
                "FTIR preprocessing pipelines — wavelet denoising, asPLS baseline "
                "correction, SNV normalisation — for reproducible ML workflows."
            ),
            "source": "https://link.springer.com/article/10.1557/s43580-026-01613-8",
            "impact": "high",
        },
        {
            "category": "Processing",
            "title": "CatBoost + Genetic Algorithms Reduce Injection-Moulding Cycle Time by 4.5 %",
            "description": (
                "An integrated CatBoost surrogate model with genetic algorithm optimisation "
                "demonstrated a 4.5 % cycle-time reduction in injection moulding while "
                "maintaining part quality (MDPI AI, 2025). A complementary study applied "
                "physics-informed neural networks to composite cure-cycle optimisation, "
                "cutting cure times by up to 30 % and reducing energy consumption. "
                "The Society of Plastics Engineers formalised industry training with "
                "dedicated AI-for-injection-moulding workshops in 2025, signalling that "
                "ML-driven process control is transitioning from research to standard "
                "industrial practice."
            ),
            "source": "https://www.eurekalert.org/news-releases/1116185",
            "impact": "medium",
        },
        {
            "category": "Sustainability",
            "title": "AI-Driven Sorting Achieves 60 % Efficiency Gains and 95 % Separation Accuracy",
            "description": (
                "AI-enabled recycling systems integrating hyperspectral imaging, FTIR, "
                "Raman, and LIBS sensors with deep-learning classifiers now achieve up to "
                "95 % material-separation accuracy and report 60 % overall efficiency "
                "gains versus conventional sorting lines (Plastics Today, 2026). Nature "
                "Reviews Materials (2026) published a landmark review positioning AI as "
                "the primary enabler of closed-loop materials economies. AI-assisted "
                "kinetic modelling and enzyme design are simultaneously accelerating "
                "chemical recycling routes — pyrolysis, solvolysis, and biocatalytic "
                "depolymerisation — toward commercial scale."
            ),
            "source": "https://www.nature.com/articles/s41578-026-00912-8",
            "impact": "high",
        },
        {
            "category": "Generative Design",
            "title": "polyGT Decoder-Only Transformer Unifies Polymer Comprehension and Inverse Design",
            "description": (
                "The polyGT model (Chemical Engineering Science, 2025) is a GPT-style "
                "decoder-only transformer trained on ~18,000 polymer SMILES that supports "
                "both property comprehension and target-driven inverse design within a "
                "single architecture. Unlike encoder-only models such as polyBERT, polyGT "
                "generates novel repeat-unit sequences conditioned on multiple simultaneous "
                "property targets — dielectric constant, thermal stability, and mechanical "
                "strength — making it suitable for multi-objective polymer discovery. "
                "The model represents a practical step toward a unified polymer foundation model."
            ),
            "source": "https://www.sciencedirect.com/science/article/abs/pii/S0009250925020159",
            "impact": "medium",
        },
        {
            "category": "Sustainability",
            "title": "AI Inverse Design of Recyclable Vitrimeric Polymers (Advanced Science 2025)",
            "description": (
                "A 2025 Advanced Science study combined MD simulation with ML surrogate "
                "models to inversely design vitrimers with targeted glass-transition "
                "temperatures and exchange-reaction kinetics. Synthesised candidates "
                "demonstrated both room-temperature healability and high-temperature "
                "flowability — enabling thermoset-equivalent mechanical performance with "
                "thermoplastic-like recyclability. This work is rare in treating "
                "recyclability as a primary design objective rather than a post-hoc "
                "consideration, offering a blueprint for the next generation of "
                "sustainable high-performance polymers."
            ),
            "source": "https://advanced.onlinelibrary.wiley.com/doi/10.1002/advs.202411385",
            "impact": "medium",
        },
        {
            "category": "LLMs",
            "title": "Fine-Tuned GPT-3.5 Classifies Polymer Solubility Across 3 373 Solvent-Polymer Pairs",
            "description": (
                "Researchers fine-tuned GPT-3.5 on 3 373 polymer–solvent pairs, achieving "
                "90 % accuracy for soluble and 83 % for insoluble classifications — "
                "comparable to purpose-built QSPR models (ACS Materials Letters, 2025). "
                "The study demonstrated that prompt engineering and chain-of-thought "
                "reasoning further boost accuracy without additional fine-tuning, offering "
                "a low-cost route to polymer informatics for organisations without "
                "dedicated ML infrastructure. LLM approaches are particularly impactful "
                "for sparse datasets where traditional QSPR struggles with extrapolation."
            ),
            "source": "https://pubs.acs.org/doi/full/10.1021/acsmaterialslett.5c00054",
            "impact": "medium",
        },
        {
            "category": "Property Prediction",
            "title": "POINT2 and OpenPoly: Open Benchmarks Attack Polymer Data-Scarcity Bottleneck",
            "description": (
                "POINT2 (POlymer INformatics Training and Testing, arXiv March 2025) "
                "provides a standardised benchmark protocol covering diverse property "
                "prediction tasks, enabling rigorous model comparison across the community "
                "for the first time. OpenPoly contributes 3 985 experimentally validated "
                "data points spanning 26 key properties derived from automated literature "
                "mining (Chinese Journal of Polymer Science, 2025). Together these "
                "resources reduce the data-scarcity bottleneck and support reproducible "
                "benchmarking — prerequisites for scientific credibility and industrial "
                "adoption of polymer ML models."
            ),
            "source": "https://arxiv.org/abs/2503.23491",
            "impact": "medium",
        },
    ],

    "key_trends": [
        "Generative AI transitions from property prediction to lab-validated inverse design of novel polymers",
        "Multimodal LLMs (text + molecular structure) outperform single-modality models in low-data regimes",
        "AI-driven recycling sorting achieves >95 % accuracy, making circular-economy targets operationally achievable",
        "Open benchmark databases (POINT2, OpenPoly) standardise evaluation and accelerate reproducible polymer informatics",
        "Closed-loop ML process control for recycled feedstocks becomes a production-readiness requirement across industry",
    ],

    "notable_tools": [
        {
            "name": "POLYT5 / polyBART",
            "purpose": (
                "Georgia Tech chemical-language transformers that generate chemically valid "
                "polymer SMILES from property targets — the first lab-validated generative "
                "polymer design AI (March 2026)."
            ),
            "url": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
        },
        {
            "name": "PolyLLMem",
            "purpose": (
                "Multimodal property-prediction model fusing Llama 3 text embeddings with "
                "Uni-Mol structural embeddings via LoRA; excels in low-data polymer regimes "
                "(Chemistry of Materials, 2025)."
            ),
            "url": "https://pubs.acs.org/doi/10.1021/acs.chemmater.5c00940",
        },
        {
            "name": "polyGT",
            "purpose": (
                "GPT-style decoder-only transformer for both polymer SMILES comprehension "
                "and multi-property inverse design from a single unified architecture "
                "(Chemical Engineering Science, 2025)."
            ),
            "url": "https://www.sciencedirect.com/science/article/abs/pii/S0009250925020159",
        },
        {
            "name": "MatterChat",
            "purpose": (
                "Multimodal LLM framework (Nature Machine Intelligence, 2026) integrating "
                "structural data with language models for high-precision, interpretable "
                "materials property predictions."
            ),
            "url": "https://www.nature.com/articles/s42256-026-01214-y",
        },
        {
            "name": "OpenPoly Database",
            "purpose": (
                "Open curated database of 3 985 experimentally validated polymer-property "
                "data points across 26 properties, enabling multi-property ML training and "
                "benchmarking."
            ),
            "url": "https://link.springer.com/article/10.1007/s10118-025-3402-y",
        },
        {
            "name": "xpectrass",
            "purpose": (
                "Open-source Python framework for systematic FTIR preprocessing evaluation "
                "(wavelet denoising, asPLS baseline, SNV normalisation) and interpretable "
                "ML polymer classification."
            ),
            "url": "https://chemrxiv.org/doi/full/10.26434/chemrxiv.15000568/v1",
        },
    ],

    "outlook": (
        "The convergence of generative AI, multimodal LLMs, and robust open benchmark "
        "databases is set to compress the polymer R&D cycle from years to weeks. As "
        "POLYT5-class models mature and are integrated with high-throughput synthesis "
        "platforms, autonomous design-make-test loops for functional polymers will move "
        "from academic demonstration to routine industrial practice. Sustainability will "
        "remain the dominant commercial driver: EU and US regulatory pressure is "
        "accelerating adoption of AI-powered sorting and chemical-recycling optimisation, "
        "and the AI Circular Economy Conference (Cologne, 2026) signalled that circularity "
        "is now an operational imperative rather than a marketing claim. Benchmark "
        "standardisation via POINT2 and OpenPoly will attract greater investment from "
        "chemical majors seeking credible, reproducible ML infrastructure. The remaining "
        "frontier — combining molecular-scale generative design with meso-scale processing "
        "simulation under a single differentiable pipeline — is expected to define the "
        "next wave of breakthrough publications and industry partnerships through 2027."
    ),
}

# ── Generate documents ─────────────────────────────────────────────────────

output_dir = Path(__file__).parent / "output"
output_dir.mkdir(exist_ok=True)

date_str = data["date"]
docx_path = output_dir / f"Polymer_AI_Digest_{date_str}.docx"
pptx_path = output_dir / f"Polymer_AI_Digest_{date_str}.pptx"

print("Generating Word document ...")
generate_word(data, docx_path)

print("Generating PowerPoint ...")
generate_pptx(data, pptx_path)

print(f"\nDone!  Files written to {output_dir.resolve()}")
