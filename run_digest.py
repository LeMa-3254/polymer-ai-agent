#!/usr/bin/env python3
"""Weekly polymer AI digest generator — 2026-05-22"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from agent import generate_word, generate_pptx

DATE = "2026-05-22"

data = {
    "date": DATE,
    "headline": (
        "From lab bench to language model: generative AI produces its first "
        "experimentally validated polymer in 2026."
    ),
    "executive_summary": (
        "May 2026 marks a pivotal moment for AI-driven polymer science. Georgia Tech's "
        "Ramprasad Group unveiled POLYT5, the first encoder-decoder foundation chemical "
        "language model for generative polymer design to be experimentally validated in the "
        "laboratory. Trained on over 100 million polymer structures in SELFIES notation, "
        "POLYT5 proposed a novel dielectric material meeting demanding targets for dielectric "
        "constant, bandgap, and glass transition temperature. That candidate was subsequently "
        "synthesised and confirmed in the wet lab, closing the loop between in-silico "
        "generation and physical reality for the first time in polymer AI history.\n\n"
        "Property prediction continues to mature rapidly. Hybrid ML frameworks combining "
        "composition and sequence features now achieve R² > 0.95 on multi-property "
        "composite datasets across density, moduli, impact strength, and thermal conductivity. "
        "Large language models (LLMs) have entered the polymer prediction arena, with "
        "multimodal architectures fusing chemical language embeddings with structural "
        "descriptors setting new accuracy records on solubility and thermal-property tasks, "
        "while LLM-powered data extraction is automatically mining polymer property records "
        "from millions of journal articles to alleviate the chronic data-scarcity bottleneck.\n\n"
        "Characterisation and processing are being reshaped by deep learning. A 2026 MRS "
        "Advances study demonstrated that neural networks predict FTIR spectra in the polymer "
        "fingerprint region with up to 87% accuracy, and machine learning now outperforms "
        "human experts on microplastic ATR-FTIR classification while surfacing labelling "
        "errors in existing datasets. On the factory floor, AI-integrated injection-moulding "
        "platforms and self-optimising composite-manufacturing pipelines are translating "
        "research gains into measurable improvements in throughput, quality, and energy use.\n\n"
        "Sustainability is the fastest-growing application domain. AI-enhanced sorting lines "
        "are reporting up to 60% efficiency improvements at commercial scale, reaching 95% "
        "material-separation accuracy for mixed-polymer streams. The new TROPIC "
        "thermodynamics database is accelerating ML-guided catalyst and enzyme design for "
        "ring-opening depolymerisation, while a landmark 2026 Nature Reviews Materials "
        "perspective formally frames AI as the critical enabler of a truly circular "
        "polymer economy."
    ),
    "findings": [
        {
            "category": "Generative Design",
            "title": "POLYT5: First Experimentally Validated Generative Polymer Foundation Model",
            "description": (
                "Georgia Tech's Ramprasad Group published POLYT5 in npj Artificial Intelligence "
                "(March 2026), an encoder-decoder T5-based model trained on 100M+ polymer "
                "structures in SELFIES notation. Deployed inside an agentic LLM framework "
                "allowing natural-language property specification, it generated 18,000+ dielectric "
                "polymer candidates and one was physically synthesised and confirmed to match "
                "predictions. This is the first end-to-end experimental validation of a "
                "generative polymer model, setting a new benchmark for the entire field."
            ),
            "source": "https://www.nature.com/articles/s44387-026-00087-1",
            "impact": "high",
        },
        {
            "category": "Generative Design",
            "title": "polyBART and polyGT: Chemical Linguists for Inverse Polymer Design",
            "description": (
                "The Ramprasad lab's polyBART (BART-based seq2seq model) enables inverse design "
                "by conditioning generation on target property values. The independently published "
                "polyGT (decoder-only transformer) unifies polymer informatics comprehension and "
                "target-driven design in a single architecture, supporting simultaneous "
                "multi-property objectives. These models collectively represent a new generation "
                "of polymer-specific language models that move decisively beyond property "
                "prediction into structure generation."
            ),
            "source": "https://www.sciencedirect.com/science/article/abs/pii/S0009250925020159",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "Hybrid ML Frameworks Achieve R² > 0.95 on Polymer Composite Properties",
            "description": (
                "A 2025 study in Materials Genome Engineering Advances trained XGBoost and "
                "ensemble models on 1,774 experimental data points for polymer composites, "
                "predicting density, HDT, flexural/tensile moduli, impact strength, and thermal "
                "conductivity with an average R² of 0.95. A JCIM companion study introduced "
                "hybrid frameworks jointly encoding composition and sequence features, outperforming "
                "single-input approaches. A 2025 uncertainty-quantification benchmark now provides "
                "practitioners with calibrated confidence intervals essential for industrial "
                "deployment."
            ),
            "source": "https://onlinelibrary.wiley.com/doi/full/10.1002/mgea.70027",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "LLMs Enter Polymer Property Prediction and Literature Mining at Scale",
            "description": (
                "ACS Materials Letters (2025) demonstrated fine-tuned GPT-3.5 achieving 90% "
                "accuracy on polymer-solvent solubility classification across 3,373 pairs. A "
                "multimodal architecture fusing LLM embeddings with structural descriptors set "
                "new benchmarks in Chemistry of Materials 2025. Simultaneously, Nature "
                "Communications-published frameworks are auto-extracting polymer-property data "
                "pairs from journal articles at scale, and a 2026 Advanced Functional Materials "
                "review benchmarked multiple LLM architectures across materials science tasks."
            ),
            "source": "https://pubs.acs.org/doi/full/10.1021/acsmaterialslett.5c00054",
            "impact": "high",
        },
        {
            "category": "Characterization",
            "title": "Deep Learning Predicts Polymer FTIR Spectra and Outperforms Human Experts",
            "description": (
                "A 2026 MRS Advances paper achieved 87% FTIR spectral prediction accuracy for "
                "polystyrene, 84% for HDPE, and 81% for PET in the polymer fingerprint region. "
                "A parallel study showed machine learning outperforms human labellers on "
                "microplastic ATR-FTIR classification while uncovering existing labelling errors. "
                "AI developments in vibrational spectroscopy highlighted in Spectroscopy Online "
                "2025 include CNN-based Raman imaging for rapid, selective polymer fingerprinting "
                "and AI-aided CEF assessment of polyolefin resins."
            ),
            "source": "https://link.springer.com/article/10.1557/s43580-026-01613-8",
            "impact": "medium",
        },
        {
            "category": "Processing",
            "title": "AI-Optimised Composite Manufacturing and Smart Injection Moulding",
            "description": (
                "A 2026 Frontiers of Chemical Science and Engineering roadmap outlined fully "
                "AI-integrated, self-optimising polymer composite production systems with "
                "real-time parameter control across heat, pressure, and flow variables. "
                "Mitsubishi Electric's Smart Injection Moulding platform uses embedded AI to "
                "monitor mould variables in real-time, targeting energy and scrap reductions. "
                "Active learning combined with explainable AI for multi-objective optimisation "
                "of spin-coated polymer films (arXiv 2025) enables Pareto-front navigation "
                "with significantly fewer experiments."
            ),
            "source": "https://journal.hep.com.cn/fcse/EN/10.1007/s11705-026-2637-7",
            "impact": "medium",
        },
        {
            "category": "Sustainability",
            "title": "AI Boosts Plastic Sorting Efficiency 60% and Accelerates Chemical Recycling",
            "description": (
                "AI-enhanced sensor sorting reported 60% efficiency gains and up to 95% "
                "material-separation accuracy for mixed-polymer streams (Plastics Today 2026). "
                "The new TROPIC database (RSC Faraday Discussions 2026) catalogues "
                "ring-opening polymerisation thermodynamics to enable ML-assisted catalyst "
                "and enzyme design for chemical recycling. A Nature Reviews Materials 2026 "
                "perspective formally positions AI as the key enabler of circular polymer "
                "economies, citing pyrolysis, solvolysis, and biocatalytic depolymerisation "
                "as priority targets."
            ),
            "source": "https://www.nature.com/articles/s41578-026-00912-8",
            "impact": "high",
        },
        {
            "category": "Generative Design",
            "title": "AI-Guided Inverse Design of Recyclable Vitrimeric Polymers",
            "description": (
                "Advanced Science (2025) published an MD-ML framework targeting vitrimers "
                "— covalent-adaptable networks prized for recyclability. The workflow "
                "performs inverse design for targeted glass transition temperatures while "
                "preserving the dynamic bond chemistry that enables reprocessability. "
                "Synthesised candidates demonstrated both room-temperature healability "
                "and high-temperature flowability, offering a blueprint for next-generation "
                "sustainable high-performance polymers that treat recyclability as a "
                "primary design objective."
            ),
            "source": "https://advanced.onlinelibrary.wiley.com/doi/10.1002/advs.202411385",
            "impact": "medium",
        },
    ],
    "key_trends": [
        "Agentic LLM wrappers around polymer foundation models enable natural-language design workflows",
        "Experimental validation loops are closing: generative models now routinely propose candidates that go to synthesis",
        "Sustainability and chemical recycling are the fastest-growing application domain for polymer AI in 2026",
        "Uncertainty quantification and explainability are becoming table-stakes requirements for industrial AI deployment",
        "Multimodal ML combining sequence, graph, spectral, and language embeddings is the dominant property-prediction paradigm",
    ],
    "notable_tools": [
        {
            "name": "POLYT5",
            "purpose": (
                "Georgia Tech encoder-decoder foundation model for generative polymer design, "
                "trained on 100M+ structures and experimentally validated on a novel "
                "dielectric material (Nature npj AI, March 2026)."
            ),
            "url": "https://www.nature.com/articles/s44387-026-00087-1",
        },
        {
            "name": "polyBART",
            "purpose": (
                "BART-based chemical language model for simultaneous polymer property "
                "prediction and inverse design via conditional sequence generation "
                "(Ramprasad Group, Georgia Tech)."
            ),
            "url": "https://arxiv.org/html/2506.04233v2",
        },
        {
            "name": "polyBERT",
            "purpose": (
                "Chemical language model for ultrafast polymer fingerprinting and "
                "informatics, outperforming handcrafted descriptors by two orders "
                "of magnitude in speed (Nature Communications 2023)."
            ),
            "url": "https://www.nature.com/articles/s41467-023-39868-6",
        },
        {
            "name": "TROPIC Database",
            "purpose": (
                "Ring-opening polymerisation thermodynamics dataset enabling "
                "ML-assisted catalyst and enzyme design for chemical recycling "
                "(RSC Faraday Discussions 2026)."
            ),
            "url": "https://pubs.rsc.org/en/content/articlehtml/2026/fd/d5fd00098j",
        },
        {
            "name": "Khazana",
            "purpose": (
                "Online platform of DFT-computed polymer property datasets "
                "(band gap, dielectric constant) with integrated ML prediction tools "
                "(Argonne ALCF / ANL)."
            ),
            "url": "https://acdc.alcf.anl.gov/mdf/detail/khazana_polymer_v1.1/",
        },
        {
            "name": "PoLyInfo",
            "purpose": (
                "Comprehensive NIMS polymer literature database covering homopolymers, "
                "copolymers, blends, and composites with structured property data "
                "for ML training."
            ),
            "url": "https://polymer.nims.go.jp/en/",
        },
    ],
    "outlook": (
        "The next 12-18 months will see agentic polymer design pipelines move from academic "
        "proof-of-concept to industrial pilot programs, particularly in high-value specialty "
        "polymers for electronics, biomedical, and sustainable packaging markets. Foundation "
        "models like POLYT5 will be extended with larger training corpora and tighter "
        "integration with automated synthesis robotics, further shrinking the "
        "discovery-to-validation cycle. LLM-powered data extraction will continue to "
        "populate polymer databases at machine speed, gradually alleviating the data "
        "scarcity that has constrained ML performance for years. Sustainability "
        "applications — AI-optimised sorting, catalytic depolymerisation, and "
        "recyclable-thermoset design — are set to attract the largest share of research "
        "funding as regulatory pressure on plastic waste intensifies globally. "
        "Practitioners should prioritise deploying uncertainty-quantified models and "
        "building explainability pipelines now, as regulatory and procurement frameworks "
        "for AI-designed materials are already being drafted in the EU and US."
    ),
}

# ── Generate documents ─────────────────────────────────────────────────────

output_dir = Path(__file__).parent / "output"
output_dir.mkdir(exist_ok=True)

docx_path = output_dir / f"Polymer_AI_Digest_{DATE}.docx"
pptx_path = output_dir / f"Polymer_AI_Digest_{DATE}.pptx"

print("Generating Word document...")
generate_word(data, docx_path)

print("Generating PowerPoint...")
generate_pptx(data, pptx_path)

print(f"\nDone! Files written to {output_dir.resolve()}")
print(f"  {docx_path.name}")
print(f"  {pptx_path.name}")
